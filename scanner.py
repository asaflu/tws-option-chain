"""
TWS Option Scanner — "Interesting" Sheet (READ-ONLY, no trading)

Monitors multiple tickers from the "interesting" sheet in twschain.xlsx,
scans their option chains for contracts that had explosive moves (+200%)
and then corrected (~50% drop from peak). Persists price history in SQLite
for multi-day tracking and logs signals for performance evaluation.

Runs as a separate process alongside option_chain.py with clientId=11.

SAFETY: This script ONLY reads market data. It does NOT place, modify,
or cancel any orders. No trading activity whatsoever.
"""

import io
import logging
import re
import sqlite3
import sys
import time
from datetime import datetime, timedelta

from ib_insync import IB, Stock, Option, util
from workbook import open_workbook, fmt_expiry

# Suppress noisy TWS error messages
logging.getLogger("ib_insync.wrapper").setLevel(logging.CRITICAL)

# ── Configuration ──────────────────────────────────────────────────
TWS_HOST = "127.0.0.1"
TWS_PORT = 7496
CLIENT_ID = 11
SCAN_INTERVAL = 60  # seconds between full scans
DB_PATH = "scanner.db"

# Defaults (overridden by sheet values)
DEFAULT_SURGE_PCT = 200
DEFAULT_DROP_PCT = 50

# Strike filtering for scanner: 5%-50% OTM range
MIN_OPTION_PRICE = 0.10  # Ignore options worth less than this
MIN_DELTA = 0.05  # Ignore options with |delta| below this
OTM_MIN_PCT = 0.05
OTM_MAX_PCT = 0.50
MAX_EXPIRY_DAYS = 7

# Snapshot pacing
BATCH_SIZE = 10
BATCH_DELAY = 1.0  # seconds between batches


# ── TWS connection ────────────────────────────────────────────────

def _on_error(*args):
    """Suppress noisy 'Unknown contract' (200) errors."""
    if len(args) >= 2 and args[1] == 200:
        return
    if len(args) >= 3:
        print(f"  TWS Error {args[1]}: {args[2]}")


def connect_tws():
    """Connect to TWS with clientId=11. Read-only data connection."""
    ib = IB()
    ib.connect(TWS_HOST, TWS_PORT, clientId=CLIENT_ID, readonly=True)
    ib.errorEvent += _on_error
    print(f"Connected to TWS at {TWS_HOST}:{TWS_PORT} (clientId={CLIENT_ID}, read-only)")
    return ib


def _qualify_quietly(ib, *contracts):
    """Qualify contracts while suppressing ib_insync stdout spam."""
    old_stdout = sys.stdout
    sys.stdout = io.StringIO()
    try:
        ib.qualifyContracts(*contracts)
    finally:
        sys.stdout = old_stdout


# ── SQLite ─────────────────────────────────────────────────────────

def init_db():
    """Create SQLite tables if they don't exist."""
    db = sqlite3.connect(DB_PATH)
    db.execute("PRAGMA journal_mode=WAL")
    db.executescript("""
        CREATE TABLE IF NOT EXISTS price_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            contract_id INTEGER NOT NULL,
            symbol TEXT NOT NULL,
            expiry TEXT NOT NULL,
            strike REAL NOT NULL,
            right TEXT NOT NULL,
            timestamp TEXT NOT NULL,
            last_price REAL,
            bid REAL,
            ask REAL,
            volume REAL,
            open_interest REAL,
            session_high REAL,
            prev_close REAL
        );
        CREATE INDEX IF NOT EXISTS idx_ph_contract
            ON price_history(contract_id);
        CREATE INDEX IF NOT EXISTS idx_ph_symbol
            ON price_history(symbol);
        CREATE INDEX IF NOT EXISTS idx_ph_contract_ts
            ON price_history(contract_id, timestamp);

        CREATE TABLE IF NOT EXISTS signals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            signal_time TEXT NOT NULL,
            symbol TEXT NOT NULL,
            expiry TEXT NOT NULL,
            strike REAL NOT NULL,
            right TEXT NOT NULL,
            entry_price REAL,
            peak_price REAL,
            baseline_price REAL,
            surge_pct REAL,
            drop_pct REAL,
            status TEXT DEFAULT 'ACTIVE',
            last_checked TEXT,
            current_price REAL,
            pnl_pct REAL
        );
        CREATE INDEX IF NOT EXISTS idx_sig_symbol
            ON signals(symbol);
        CREATE INDEX IF NOT EXISTS idx_sig_status
            ON signals(status);
    """)
    db.commit()
    return db


# ── Excel helpers ──────────────────────────────────────────────────

def get_sheets(wb):
    """Get 'interesting' and 'signals_log' sheets."""
    return wb.sheets["interesting"], wb.sheets["signals_log"]


def read_config(ws):
    """Read tickers and thresholds from the 'interesting' sheet."""
    raw = ws.range("A2:A21").value
    tickers = []
    for val in raw:
        if val is not None and str(val).strip():
            sym = re.sub(r"[^A-Za-z]", "", str(val)).upper()
            if sym:
                tickers.append(sym)

    surge = ws.range("E1").value
    drop = ws.range("E2").value
    try:
        surge = float(surge)
    except (TypeError, ValueError):
        surge = DEFAULT_SURGE_PCT
    try:
        drop = float(drop)
    except (TypeError, ValueError):
        drop = DEFAULT_DROP_PCT

    return tickers, surge, drop


# ── TWS scanning ──────────────────────────────────────────────────

def get_stock_price_snapshot(ib, stock):
    """Get current stock price via snapshot. Returns price or None."""
    ticker = ib.reqMktData(stock, "", True, False)
    ib.sleep(2)
    price = ticker.marketPrice()
    if util.isNan(price):
        price = ticker.close
    if util.isNan(price):
        return None
    return price


def get_option_params(ib, stock):
    """Get option chain parameters for expiries within MAX_EXPIRY_DAYS.
    Returns list of (expiry, strikes, tradingClass)."""
    chains = ib.reqSecDefOptParams(stock.symbol, "", stock.secType, stock.conId)
    if not chains:
        return []

    smart_chains = [c for c in chains if c.exchange == "SMART"]
    if not smart_chains:
        smart_chains = chains

    chain = next((c for c in smart_chains if c.tradingClass == stock.symbol), smart_chains[0])

    expirations = sorted(chain.expirations)
    today = datetime.now().strftime("%Y%m%d")
    max_date = (datetime.now() + timedelta(days=MAX_EXPIRY_DAYS)).strftime("%Y%m%d")
    near_exps = [e for e in expirations if today <= e <= max_date]

    if not near_exps:
        return []

    return [(exp, sorted(chain.strikes), chain.tradingClass) for exp in near_exps]


def filter_otm_strikes(all_strikes, stock_price):
    """Filter to OTM strikes in the 5%-50% range from stock price.
    Returns (call_strikes, put_strikes) — calls above price, puts below."""
    call_lo = stock_price * (1 + OTM_MIN_PCT)
    call_hi = stock_price * (1 + OTM_MAX_PCT)
    put_lo = stock_price * (1 - OTM_MAX_PCT)
    put_hi = stock_price * (1 - OTM_MIN_PCT)

    # Also include near-ATM (within 3%)
    atm_lo = stock_price * 0.97
    atm_hi = stock_price * 1.03

    call_strikes = [s for s in all_strikes if (call_lo <= s <= call_hi) or (atm_lo <= s <= atm_hi)]
    put_strikes = [s for s in all_strikes if (put_lo <= s <= put_hi) or (atm_lo <= s <= atm_hi)]

    return call_strikes, put_strikes


def get_snapshots(ib, contracts):
    """Request snapshots in batches with pacing. Returns list of (contract, ticker)."""
    results = []
    for i in range(0, len(contracts), BATCH_SIZE):
        batch = contracts[i:i + BATCH_SIZE]
        for c in batch:
            t = ib.reqMktData(c, "", True, False)  # snapshot=True
            results.append((c, t))
        ib.sleep(BATCH_DELAY)
    # Give final batch time to settle
    if contracts:
        ib.sleep(1)
    return results


def _snap_val(v):
    """Extract numeric value from ticker field, return None if NaN."""
    if v is None or (isinstance(v, float) and util.isNan(v)):
        return None
    return float(v)


def _get_mid_price(last, bid, ask):
    """Derive best available price from snapshot fields."""
    if bid is not None and ask is not None:
        return (bid + ask) / 2
    if last is not None:
        return last
    if bid is not None:
        return bid
    if ask is not None:
        return ask
    return None


def scan_ticker(ib, symbol, db):
    """Full scan of one ticker's option chain. Returns number of prices stored."""
    stock = Stock(symbol, "SMART", "USD")
    _qualify_quietly(ib, stock)

    if not stock.conId:
        print(f"  Could not qualify stock '{symbol}'")
        return 0

    stock_price = get_stock_price_snapshot(ib, stock)
    if stock_price is None:
        print(f"  No price for '{symbol}'")
        return 0

    print(f"  {symbol} ${stock_price:.2f}")

    params = get_option_params(ib, stock)
    if not params:
        print(f"  No near-term options for '{symbol}'")
        return 0

    now = datetime.now().isoformat()
    total_stored = 0

    for expiry, all_strikes, trading_class in params:
        call_strikes, put_strikes = filter_otm_strikes(all_strikes, stock_price)

        contracts = []
        for strike in call_strikes:
            contracts.append(Option(symbol, expiry, strike, "C", "SMART",
                                    tradingClass=trading_class))
        for strike in put_strikes:
            contracts.append(Option(symbol, expiry, strike, "P", "SMART",
                                    tradingClass=trading_class))

        if not contracts:
            continue

        _qualify_quietly(ib, *contracts)
        valid = [c for c in contracts if c.conId]
        if not valid:
            continue

        print(f"    {fmt_expiry(expiry)}: {len(valid)} contracts")

        snapshots = get_snapshots(ib, valid)
        total_stored += store_prices(db, snapshots, now)

    return total_stored


def store_prices(db, snapshots, timestamp):
    """Persist snapshot data to price_history table. Returns count stored."""
    rows = []
    for contract, ticker in snapshots:
        last = _snap_val(ticker.last)
        bid = _snap_val(ticker.bid)
        ask = _snap_val(ticker.ask)

        # Skip if we have no price data at all
        if last is None and bid is None and ask is None:
            continue

        # Skip options below minimum price threshold
        mid = _get_mid_price(last, bid, ask)
        if mid is not None and mid < MIN_OPTION_PRICE:
            continue

        # Skip options with delta below minimum
        greeks = ticker.modelGreeks
        if greeks is not None:
            delta = greeks.delta
            if delta is not None and not util.isNan(delta) and abs(delta) < MIN_DELTA:
                continue

        vol = _snap_val(ticker.volume)
        oi = _snap_val(ticker.callOpenInterest if contract.right == "C"
                       else ticker.putOpenInterest)
        high = _snap_val(ticker.high)
        prev_close = _snap_val(ticker.close)

        rows.append((
            contract.conId, contract.symbol, contract.lastTradeDateOrContractMonth,
            contract.strike, contract.right, timestamp,
            last, bid, ask, vol, oi, high, prev_close
        ))

    if rows:
        db.executemany("""
            INSERT INTO price_history
            (contract_id, symbol, expiry, strike, right, timestamp,
             last_price, bid, ask, volume, open_interest, session_high, prev_close)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)
        # Don't commit here — let the caller batch commits per scan cycle

    return len(rows)


# ── Signal analysis ───────────────────────────────────────────────

def analyze_signals(db, symbol, surge_threshold, drop_threshold):
    """Find interesting contracts for a symbol. Returns list of signal dicts."""
    candidates = db.execute("""
        SELECT contract_id, expiry, strike, right, COUNT(*) as cnt
        FROM price_history
        WHERE symbol = ?
        GROUP BY contract_id
        HAVING cnt >= 2
    """, (symbol,)).fetchall()

    signals = []

    for contract_id, expiry, strike, right, cnt in candidates:
        rows = db.execute("""
            SELECT last_price, bid, ask
            FROM price_history
            WHERE contract_id = ?
            ORDER BY timestamp
        """, (contract_id,)).fetchall()

        prices = []
        for last, bid, ask in rows:
            price = _get_mid_price(last, bid, ask)
            if price is not None and price > 0:
                prices.append(price)

        if len(prices) < 2:
            continue

        # Baseline: minimum of first 25% of observations
        first_quarter = max(1, len(prices) // 4)
        baseline = min(prices[:first_quarter])

        if baseline <= 0:
            continue

        peak = max(prices)
        current = prices[-1]

        surge_pct = (peak - baseline) / baseline * 100
        drop_pct = (peak - current) / peak * 100 if peak > 0 else 0

        if surge_pct >= surge_threshold and drop_pct >= drop_threshold:
            signals.append({
                "contract_id": contract_id,
                "symbol": symbol,
                "expiry": expiry,
                "strike": strike,
                "right": right,
                "current": current,
                "peak": peak,
                "baseline": baseline,
                "surge_pct": round(surge_pct, 1),
                "drop_pct": round(drop_pct, 1),
                "score": surge_pct * drop_pct,
            })

    # Sort by interestingness, top 2 per ticker
    signals.sort(key=lambda s: s["score"], reverse=True)
    return signals[:2]


# ── Writing results ───────────────────────────────────────────────

def write_results(ws, all_signals):
    """Write signal data to the 'interesting' sheet signal table (J2+)."""
    ws.range("H1").value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.range("J2:S51").clear_contents()

    if not all_signals:
        return

    rows = []
    for sig in all_signals:
        rows.append([
            sig["symbol"],
            fmt_expiry(sig["expiry"]),
            sig["strike"],
            sig["right"],
            sig["current"],
            sig["peak"],
            sig["baseline"],
            sig["surge_pct"],
            sig["drop_pct"],
            sig.get("signal_time", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ])

    ws.range("J2").value = rows


def update_signal_log(db, ws_log, all_signals):
    """Log new signals to DB and update signals_log sheet."""
    now = datetime.now().isoformat()

    for sig in all_signals:
        existing = db.execute("""
            SELECT id FROM signals
            WHERE symbol = ? AND expiry = ? AND strike = ? AND right = ?
              AND status = 'ACTIVE'
        """, (sig["symbol"], sig["expiry"], sig["strike"], sig["right"])).fetchone()

        if existing:
            db.execute("""
                UPDATE signals
                SET current_price = ?, last_checked = ?,
                    pnl_pct = CASE WHEN entry_price > 0
                              THEN ROUND((? - entry_price) / entry_price * 100, 1)
                              ELSE 0 END
                WHERE id = ?
            """, (sig["current"], now, sig["current"], existing[0]))
        else:
            sig["signal_time"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            db.execute("""
                INSERT INTO signals
                (signal_time, symbol, expiry, strike, right, entry_price,
                 peak_price, baseline_price, surge_pct, drop_pct,
                 status, last_checked, current_price, pnl_pct)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE', ?, ?, 0)
            """, (sig["signal_time"], sig["symbol"], sig["expiry"],
                  sig["strike"], sig["right"], sig["current"],
                  sig["peak"], sig["baseline"], sig["surge_pct"],
                  sig["drop_pct"], now, sig["current"]))

    # Lifecycle: ACTIVE → EXPIRED → WINNER/LOSER
    today = datetime.now().strftime("%Y%m%d")
    db.execute("UPDATE signals SET status = 'EXPIRED' WHERE status = 'ACTIVE' AND expiry < ?", (today,))
    db.execute("UPDATE signals SET status = 'WINNER' WHERE status = 'EXPIRED' AND pnl_pct > 0")
    db.execute("UPDATE signals SET status = 'LOSER' WHERE status = 'EXPIRED' AND pnl_pct <= 0")

    db.commit()

    # Write full signal log to sheet
    all_db_signals = db.execute("""
        SELECT signal_time, symbol, expiry, strike, right,
               entry_price, peak_price, current_price, pnl_pct, status
        FROM signals
        ORDER BY signal_time DESC
        LIMIT 100
    """).fetchall()

    if all_db_signals:
        ws_log.range("A2:J101").clear_contents()
        display_rows = []
        for row in all_db_signals:
            r = list(row)
            if r[2]:
                try:
                    r[2] = fmt_expiry(r[2])
                except ValueError:
                    pass  # Already formatted or invalid — leave as-is
            display_rows.append(r)
        ws_log.range("A2").value = display_rows


# ── Main loop ─────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  TWS Option Scanner (READ-ONLY)")
    print("  No trading activity — market data only")
    print("=" * 60)

    ib = connect_tws()
    db = init_db()
    print(f"Database: {DB_PATH}")

    wb = open_workbook()
    ws, ws_log = get_sheets(wb)
    print(f"Excel: twschain.xlsx — sheets 'interesting' and 'signals_log' ready")
    print(f"Add tickers in column A (A2:A21). Scanning every {SCAN_INTERVAL}s.\n")

    try:
        while True:
            cycle_start = time.time()

            if not ib.isConnected():
                print("Connection lost. Reconnecting in 5 seconds...")
                time.sleep(5)
                try:
                    ib.disconnect()
                    ib = connect_tws()
                except Exception as e:
                    print(f"  Reconnect failed: {e}")
                    time.sleep(SCAN_INTERVAL)
                    continue

            tickers, surge_threshold, drop_threshold = read_config(ws)

            if not tickers:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] No tickers in A2:A21. Waiting...")
                time.sleep(SCAN_INTERVAL)
                continue

            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Scanning {len(tickers)} tickers "
                  f"(surge >= {surge_threshold}%, drop >= {drop_threshold}%)")

            all_signals = []
            for symbol in tickers:
                try:
                    count = scan_ticker(ib, symbol, db)
                    print(f"    Stored {count} prices")

                    sigs = analyze_signals(db, symbol, surge_threshold, drop_threshold)
                    if sigs:
                        print(f"    ** {len(sigs)} signal(s) found!")
                        for s in sigs:
                            print(f"       {s['right']} {s['strike']} {fmt_expiry(s['expiry'])}: "
                                  f"surge {s['surge_pct']}% drop {s['drop_pct']}%")
                    all_signals.extend(sigs)

                except (ConnectionError, OSError) as e:
                    print(f"  Connection error scanning {symbol}: {e}")
                    break  # Don't continue scanning if TWS connection died
                except Exception as e:
                    print(f"  Error scanning {symbol}: {e}")

            # Commit all price_history inserts from this cycle in one batch
            try:
                db.commit()
            except sqlite3.Error as e:
                print(f"  DB commit error: {e}")

            # Write results to Excel (wrapped for resilience)
            try:
                write_results(ws, all_signals)
                update_signal_log(db, ws_log, all_signals)
                print(f"  Wrote {len(all_signals)} signals to Excel")
            except Exception as e:
                print(f"  Error writing to Excel: {e}")

            elapsed = time.time() - cycle_start
            sleep_time = max(1, SCAN_INTERVAL - elapsed)
            print(f"  Cycle took {elapsed:.0f}s, sleeping {sleep_time:.0f}s")
            time.sleep(sleep_time)

    except KeyboardInterrupt:
        print("\n\nStopping...")
    finally:
        db.close()
        ib.disconnect()
        print("Disconnected from TWS.")


if __name__ == "__main__":
    main()
