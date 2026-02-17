"""
TWS Option Chain Streamer → Excel (READ-ONLY, no trading)

Connects to TWS on port 7496 and streams the nearest-expiry option chain
for the ticker in cell A1 into Excel every 10 seconds.

SAFETY: This script ONLY reads market data. It does NOT place, modify,
or cancel any orders. No trading activity whatsoever.
"""

import io
import logging
import re
import sys
import time
from datetime import datetime

from ib_insync import IB, Stock, Option, util
import xlwings as xw

# Suppress noisy TWS error messages (e.g., Error 200 "No security definition")
logging.getLogger("ib_insync.wrapper").setLevel(logging.CRITICAL)

# ── Configuration ──────────────────────────────────────────────────
TWS_HOST = "127.0.0.1"
TWS_PORT = 7496
CLIENT_ID = 10
REFRESH_SECONDS = 10
MAX_STRIKES = 80  # Max strikes to display after pruning
MAX_TWS_LINES = 100  # TWS concurrent market data line limit
STRIKE_PCT_RANGE = 0.20  # +/- 20% of stock price


def _on_error(*args):
    """Custom error handler — suppress noisy 'Unknown contract' (200) errors."""
    # args: (reqId, errorCode, errorString) or (reqId, errorCode, errorString, contract)
    if len(args) >= 2 and args[1] == 200:
        return
    if len(args) >= 3:
        print(f"  TWS Error {args[1]}: {args[2]}")


def connect_tws():
    """Connect to TWS. Read-only data connection."""
    ib = IB()
    ib.connect(TWS_HOST, TWS_PORT, clientId=CLIENT_ID, readonly=True)
    ib.errorEvent += _on_error
    print(f"Connected to TWS at {TWS_HOST}:{TWS_PORT} (read-only)")
    return ib


def subscribe_stock(ib, stock):
    """Subscribe to streaming stock price. Returns the ticker object (kept alive)."""
    ticker = ib.reqMktData(stock, "", False, False)
    ib.sleep(2)
    return ticker


def get_stock_price(ticker):
    """Read current price from a live streaming ticker."""
    price = ticker.marketPrice()
    if util.isNan(price):
        price = ticker.close
    if util.isNan(price):
        return None
    return price


def get_nearest_expiry(ib, stock):
    """Get the nearest option expiration date for a stock.
    Returns (expiry, strikes, tradingClass) — tradingClass needed to avoid
    ambiguous contracts when multiple option classes exist (e.g. PLTR vs 2PLTR)."""
    chains = ib.reqSecDefOptParams(stock.symbol, "", stock.secType, stock.conId)
    if not chains:
        return None, None, None

    # Prefer SMART exchange, and pick the chain whose tradingClass matches the symbol
    # (filters out adjusted/mini options like 2PLTR, 7TSLA, etc.)
    smart_chains = [c for c in chains if c.exchange == "SMART"]
    if not smart_chains:
        smart_chains = chains

    # Prefer the chain with tradingClass == stock symbol (standard options)
    chain = next((c for c in smart_chains if c.tradingClass == stock.symbol), smart_chains[0])
    trading_class = chain.tradingClass

    expirations = sorted(chain.expirations)
    today = datetime.now().strftime("%Y%m%d")
    future_exps = [e for e in expirations if e >= today]

    if not future_exps:
        return None, None, None

    nearest = future_exps[0]
    print(f"  Using trading class: {trading_class}")
    return nearest, sorted(chain.strikes), trading_class


def filter_strikes(all_strikes, stock_price):
    """Filter strikes to those within STRIKE_PCT_RANGE of stock price,
    capped at MAX_STRIKES centered around ATM."""
    low = stock_price * (1 - STRIKE_PCT_RANGE)
    high = stock_price * (1 + STRIKE_PCT_RANGE)
    in_range = [s for s in all_strikes if low <= s <= high]

    if len(in_range) <= MAX_STRIKES:
        return in_range

    # Find the ATM strike and take MAX_STRIKES centered around it
    atm_idx = min(range(len(in_range)), key=lambda i: abs(in_range[i] - stock_price))
    half = MAX_STRIKES // 2
    start = max(0, atm_idx - half)
    end = start + MAX_STRIKES
    if end > len(in_range):
        end = len(in_range)
        start = max(0, end - MAX_STRIKES)
    return in_range[start:end]


def subscribe_option_chain(ib, symbol, expiry, strikes, ticker_cache, trading_class):
    """Subscribe to streaming market data for all strikes in range.
    Uses streaming subscriptions (NOT snapshots) — no extra data fees.
    No pruning — all qualified contracts stay subscribed for a complete chain."""
    contracts = []
    for strike in strikes:
        c = Option(symbol, expiry, strike, "C", "SMART", tradingClass=trading_class)
        p = Option(symbol, expiry, strike, "P", "SMART", tradingClass=trading_class)
        contracts.append(c)
        contracts.append(p)

    # Suppress "Unknown contract" print from ib_insync
    old_stdout = sys.stdout
    sys.stdout = io.StringIO()
    try:
        ib.qualifyContracts(*contracts)
    finally:
        sys.stdout = old_stdout

    valid = [c for c in contracts if c.conId]
    print(f"  Qualified: {len(valid)} of {len(contracts)} contracts")

    if not valid:
        print("  WARNING: No contracts qualified. Check if options exist for this expiry.")
        return

    # Cap at TWS limit
    if len(valid) > MAX_TWS_LINES:
        print(f"  Capping at {MAX_TWS_LINES} to stay within TWS limit")
        valid = valid[:MAX_TWS_LINES]

    print(f"  Subscribing to {len(valid)} contracts...")

    for i, contract in enumerate(valid):
        ticker = ib.reqMktData(contract, genericTickList="106")
        ticker_cache[contract.conId] = (contract, ticker)
        if (i + 1) % 10 == 0:
            ib.sleep(0.5)

    # Wait for data to arrive
    print(f"  Waiting for data...")
    ib.sleep(8)

    has_greeks = sum(1 for _, t in ticker_cache.values() if t.modelGreeks is not None)
    print(f"  Ready: {len(ticker_cache)} contracts, {has_greeks} with greeks")



def cancel_subscriptions(ib, ticker_cache):
    """Cancel all active market data subscriptions."""
    for contract, _ticker in ticker_cache.values():
        try:
            ib.cancelMktData(contract)
        except Exception:
            pass
    ticker_cache.clear()


def read_option_data(ticker_cache):
    """Read current data from active streaming subscriptions. No new API requests."""
    calls = {}
    puts = {}

    for contract, ticker in ticker_cache.values():
        strike = contract.strike
        greeks = ticker.modelGreeks

        row = {
            "bid": _val(ticker.bid),
            "ask": _val(ticker.ask),
            "last": _val(ticker.last),
            "volume": _val(ticker.volume),
            "openInterest": _val(ticker.callOpenInterest if contract.right == "C" else ticker.putOpenInterest),
            "delta": _greek(greeks, "delta"),
            "gamma": _greek(greeks, "gamma"),
            "theta": _greek(greeks, "theta"),
        }

        if contract.right == "C":
            calls[strike] = row
        else:
            puts[strike] = row

    return calls, puts


def _val(v):
    """Return value or empty string if NaN."""
    if v is None or (isinstance(v, float) and util.isNan(v)):
        return ""
    return v


def _greek(greeks, field):
    """Extract a greek value safely."""
    if greeks is None:
        return ""
    val = getattr(greeks, field, None)
    if val is None or util.isNan(val):
        return ""
    return round(val, 4)


# Column layout
CALL_HEADERS = ["C Delta", "C Gamma", "C Theta", "C Volume", "C OI", "C Bid", "C Ask", "C Last"]
STRIKE_HEADER = ["Strike"]
PUT_HEADERS = ["P Last", "P Bid", "P Ask", "P OI", "P Volume", "P Delta", "P Gamma", "P Theta"]
ALL_HEADERS = CALL_HEADERS + STRIKE_HEADER + PUT_HEADERS


def setup_excel(ws, prev_row_count):
    """One-time Excel setup on ticker change: clear old data, write headers."""
    # Clear all old data below row 1 in one call
    if prev_row_count > 0:
        ws.range(f"A2:Q{4 + prev_row_count}").clear()
    ws.range("A4").value = ALL_HEADERS
    ws.range("A4:Q4").font.bold = True


def write_to_excel(ws, calls, puts, strikes, expiry, stock_price, prev_row_count):
    """Write option chain data to Excel. Minimizes AppleScript calls.
    Layout: A1=ticker, A2=price, A3=expiry, row 4=headers, row 5+=data.
    Returns current row count."""

    data_rows = []
    for strike in sorted(strikes):
        call = calls.get(strike, {})
        put = puts.get(strike, {})
        data_rows.append([
            call.get("delta", ""),
            call.get("gamma", ""),
            call.get("theta", ""),
            call.get("volume", ""),
            call.get("openInterest", ""),
            call.get("bid", ""),
            call.get("ask", ""),
            call.get("last", ""),
            strike,
            put.get("last", ""),
            put.get("bid", ""),
            put.get("ask", ""),
            put.get("openInterest", ""),
            put.get("volume", ""),
            put.get("delta", ""),
            put.get("gamma", ""),
            put.get("theta", ""),
        ])

    num_rows = len(data_rows)

    # Clear leftover rows from previous cycle (only if shrunk)
    if prev_row_count > num_rows:
        ws.range(f"A{5 + num_rows}:Q{4 + prev_row_count}").clear_contents()

    # Single write: price + expiry
    ws.range("A2").value = [[stock_price], [expiry]]

    # Single write: all data rows at once
    if data_rows:
        ws.range("A5").value = data_rows

    return num_rows


def main():
    print("=" * 60)
    print("  TWS Option Chain Streamer (READ-ONLY)")
    print("  No trading activity — market data only")
    print("=" * 60)

    # Connect to TWS
    ib = connect_tws()

    # Connect to Excel — open twschain.xlsx or attach if already open
    try:
        wb = xw.Book("twschain.xlsx")
    except Exception:
        wb = xw.Book()
        wb.save("twschain.xlsx")

    # Use the "chain" sheet, create it if missing
    if "chain" in [s.name for s in wb.sheets]:
        ws = wb.sheets["chain"]
    else:
        ws = wb.sheets.add("chain", after=wb.sheets[-1])

    print("Ready. Type a ticker symbol in cell A1 of the active Excel sheet.")
    print(f"Data refreshes every {REFRESH_SECONDS} seconds. Press Ctrl+C to stop.\n")

    last_ticker = None
    filtered_strikes = []
    ticker_cache = {}  # conId -> (contract, ticker)
    stock_ticker = None  # Streaming stock price subscription
    prev_row_count = 0
    stock_price = None
    expiry = None

    try:
        while True:
            # Reconnect if disconnected
            if not ib.isConnected():
                print("Connection lost. Reconnecting in 5 seconds...")
                time.sleep(5)
                try:
                    ib.disconnect()
                    ib = connect_tws()
                    last_ticker = None
                    ticker_cache = {}
                except Exception as e:
                    print(f"  Reconnect failed: {e}")
                    time.sleep(REFRESH_SECONDS)
                    continue

            # Read ticker from A1
            ticker_val = ws.range("A1").value
            if ticker_val is None or str(ticker_val).strip() == "":
                time.sleep(REFRESH_SECONDS)
                continue

            # Sanitize to ASCII letters only (handles Excel encoding quirks)
            symbol = re.sub(r"[^A-Za-z]", "", str(ticker_val)).upper()
            if not symbol:
                time.sleep(REFRESH_SECONDS)
                continue

            try:
                # If ticker changed, set up new subscriptions
                if symbol != last_ticker:
                    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] New ticker: {symbol}")

                    # Cancel old subscriptions before anything else
                    if ticker_cache:
                        cancel_subscriptions(ib, ticker_cache)

                    last_ticker = symbol

                    # Cancel old stock subscription
                    if stock_ticker is not None:
                        try:
                            ib.cancelMktData(stock_ticker.contract)
                        except Exception:
                            pass
                        stock_ticker = None

                    # Qualify the stock first
                    stock = Stock(symbol, "SMART", "USD")
                    ib.qualifyContracts(stock)

                    if not stock.conId:
                        print(f"  ERROR: Could not find stock '{symbol}'")
                        ws.range("B1").value = f"ERROR: Unknown ticker '{symbol}'"
                        time.sleep(REFRESH_SECONDS)
                        last_ticker = None
                        continue

                    # Subscribe to streaming stock price (stays alive)
                    stock_ticker = subscribe_stock(ib, stock)
                    stock_price = get_stock_price(stock_ticker)
                    if stock_price is None:
                        print(f"  ERROR: Could not get price for '{symbol}'")
                        ws.range("B1").value = f"ERROR: No price for '{symbol}'"
                        time.sleep(REFRESH_SECONDS)
                        last_ticker = None
                        continue
                    print(f"  Current price: ${stock_price:.2f}")

                    # Get nearest expiry and strikes
                    expiry, all_strikes, trading_class = get_nearest_expiry(ib, stock)
                    if not expiry:
                        print(f"  ERROR: No options found for '{symbol}'")
                        ws.range("B1").value = f"ERROR: No options for '{symbol}'"
                        time.sleep(REFRESH_SECONDS)
                        last_ticker = None
                        continue

                    # Filter strikes around ATM, capped at MAX_STRIKES
                    filtered_strikes = filter_strikes(all_strikes, stock_price)
                    print(f"  Nearest expiry: {expiry} | {len(filtered_strikes)} strikes (of {len(all_strikes)} total)")

                    # Subscribe to streaming data ONCE (no snapshots, no extra fees)
                    subscribe_option_chain(ib, symbol, expiry, filtered_strikes, ticker_cache, trading_class)

                    # One-time Excel setup: clear old data, write headers
                    setup_excel(ws, prev_row_count)
                    prev_row_count = 0

                # Let event loop process incoming data
                ib.sleep(0.1)

                # Read live stock price
                if stock_ticker is not None:
                    live_price = get_stock_price(stock_ticker)
                    if live_price is not None:
                        stock_price = live_price

                # Read data from active subscriptions (no new API requests)
                print(f"[{datetime.now().strftime('%H:%M:%S')}] {symbol} ${stock_price:.2f} ({expiry})...", end=" ")
                calls, puts = read_option_data(ticker_cache)

                # Build display list from subscribed strikes only (already pruned)
                active_strikes = set()
                for contract, _t in ticker_cache.values():
                    active_strikes.add(contract.strike)
                active_strikes = sorted(active_strikes)

                # Write to Excel — 2-3 AppleScript calls max
                prev_row_count = write_to_excel(ws, calls, puts, active_strikes, expiry, stock_price, prev_row_count)
                print(f"Done — {len(active_strikes)} strikes written.")

            except (ConnectionError, OSError) as e:
                print(f"\n  Connection error: {e}")
                last_ticker = None
                ticker_cache = {}
                stock_ticker = None
                continue

            time.sleep(REFRESH_SECONDS)

    except KeyboardInterrupt:
        print("\n\nStopping...")
    finally:
        try:
            cancel_subscriptions(ib, ticker_cache)
        except Exception:
            pass
        ib.disconnect()
        print("Disconnected from TWS.")


if __name__ == "__main__":
    main()
