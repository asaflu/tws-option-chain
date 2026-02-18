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
import subprocess
import sys
import time
from datetime import datetime

from ib_insync import IB, Stock, Option, util
from workbook import open_workbook, fmt_expiry

# Suppress noisy TWS error messages (e.g., Error 200 "No security definition")
logging.getLogger("ib_insync.wrapper").setLevel(logging.CRITICAL)

# ── Configuration ──────────────────────────────────────────────────
TWS_HOST = "127.0.0.1"
TWS_PORT = 7496
CLIENT_ID = 10
REFRESH_SECONDS = 10
MAX_TWS_LINES = 100     # TWS concurrent market data line limit
MAX_STRIKES = MAX_TWS_LINES // 2  # Each strike = 1 call + 1 put
MIN_DELTA = 0.02        # Minimum |delta| to display a strike row
MULTIPLIERS = [2, 3, 5, 7, 10, 15, 20]


# ── TWS connection & data ─────────────────────────────────────────

def _on_error(*args):
    """Custom error handler — suppress noisy 'Unknown contract' (200) errors."""
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


def _qualify_quietly(ib, *contracts):
    """Qualify contracts while suppressing ib_insync stdout spam."""
    old_stdout = sys.stdout
    sys.stdout = io.StringIO()
    try:
        ib.qualifyContracts(*contracts)
    finally:
        sys.stdout = old_stdout


def subscribe_stock(ib, stock):
    """Subscribe to streaming stock price. Returns the ticker object."""
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


def get_expiries(ib, stock, count=5):
    """Get the nearest option expiration dates for a stock.
    Returns (expiries_list, strikes, tradingClass)."""
    chains = ib.reqSecDefOptParams(stock.symbol, "", stock.secType, stock.conId)
    if not chains:
        return [], None, None

    smart_chains = [c for c in chains if c.exchange == "SMART"]
    if not smart_chains:
        smart_chains = chains

    chain = next((c for c in smart_chains if c.tradingClass == stock.symbol), smart_chains[0])
    trading_class = chain.tradingClass

    expirations = sorted(chain.expirations)
    today = datetime.now().strftime("%Y%m%d")
    future_exps = [e for e in expirations if e >= today]

    if not future_exps:
        return [], None, None

    print(f"  Using trading class: {trading_class}")
    return future_exps[:count], sorted(chain.strikes), trading_class


# ── Expiry selector ───────────────────────────────────────────────

def _read_expiry_choice(ws):
    """Read the numeric expiry choice (1-5) from B3. Returns int or None."""
    val = ws.range("B3").value
    if val is None:
        return None
    try:
        choice = int(float(val))
        return choice if choice >= 1 else None
    except (ValueError, TypeError):
        return None


def setup_expiry_selector(ws, expiries):
    """Show numbered expiry options in row 3. User types 1-5 in B3 to choose."""
    labels = [f"{i+1}: {fmt_expiry(e)}" for i, e in enumerate(expiries)]
    ws.range("A3").value = "Expiry #"
    ws.range("A3").font.bold = True
    ws.range("B3").value = 1
    ws.range("C3:H3").clear_contents()
    ws.range("C3").value = "  ".join(labels)


# ── Strike filtering ──────────────────────────────────────────────

def filter_strikes(all_strikes, stock_price):
    """Filter strikes symmetrically around ATM: equal count above and below,
    capped at MAX_STRIKES total."""
    if not all_strikes:
        return []

    atm_idx = min(range(len(all_strikes)), key=lambda i: abs(all_strikes[i] - stock_price))
    available_below = atm_idx
    available_above = len(all_strikes) - atm_idx - 1
    half = min(available_below, available_above, MAX_STRIKES // 2)

    start = atm_idx - half
    end = atm_idx + half + 1
    return all_strikes[start:end]


def _passes_delta_filter(call_data, put_data):
    """Return True if both sides with delta data meet the MIN_DELTA threshold.
    A strike is excluded if ANY side has |delta| < MIN_DELTA."""
    c_delta = call_data.get("delta", "")
    p_delta = put_data.get("delta", "")
    # Need at least one side with delta data
    if c_delta == "" and p_delta == "":
        return False
    # Every side that HAS delta must pass the threshold
    if c_delta != "" and abs(c_delta) < MIN_DELTA:
        return False
    if p_delta != "" and abs(p_delta) < MIN_DELTA:
        return False
    return True


def _enforce_symmetry(strikes, stock_price):
    """Given a sorted list of strikes, return a symmetric subset around ATM."""
    if not strikes:
        return strikes
    atm_idx = min(range(len(strikes)), key=lambda i: abs(strikes[i] - stock_price))
    below = atm_idx
    above = len(strikes) - atm_idx - 1
    half = min(below, above)
    return strikes[atm_idx - half : atm_idx + half + 1]


def _get_display_strikes(calls, puts, all_strikes, stock_price):
    """Delta-filter and symmetry-enforce strikes for display. Shared by chain + 10x."""
    filtered = []
    for strike in sorted(all_strikes):
        call = calls.get(strike, {})
        put = puts.get(strike, {})
        if _passes_delta_filter(call, put):
            filtered.append(strike)
    return _enforce_symmetry(filtered, stock_price)


# ── Subscriptions ─────────────────────────────────────────────────

def subscribe_option_chain(ib, symbol, expiry, strikes, ticker_cache, trading_class):
    """Subscribe to streaming market data for all strikes in range."""
    contracts = []
    for strike in strikes:
        contracts.append(Option(symbol, expiry, strike, "C", "SMART", tradingClass=trading_class))
        contracts.append(Option(symbol, expiry, strike, "P", "SMART", tradingClass=trading_class))

    _qualify_quietly(ib, *contracts)

    valid = [c for c in contracts if c.conId]
    print(f"  Qualified: {len(valid)} of {len(contracts)} contracts")

    if not valid:
        print("  WARNING: No contracts qualified. Check if options exist for this expiry.")
        return

    # MAX_STRIKES already ensures len(valid) <= MAX_TWS_LINES,
    # but guard against edge cases (e.g. extra strikes from chain params).
    if len(valid) > MAX_TWS_LINES:
        # Trim symmetrically: drop from both ends to keep ATM centered.
        excess = len(valid) - MAX_TWS_LINES
        trim_front = (excess // 2) & ~1  # round down to even (keep C/P pairs)
        trim_back = excess - trim_front
        valid = valid[trim_front : len(valid) - trim_back]
        print(f"  Capped to {len(valid)} contracts (symmetric trim)")

    print(f"  Subscribing to {len(valid)} contracts...")

    for i, contract in enumerate(valid):
        ticker = ib.reqMktData(contract, genericTickList="106")
        ticker_cache[contract.conId] = (contract, ticker)
        if (i + 1) % 10 == 0:
            ib.sleep(0.5)

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


# ── Reading market data ───────────────────────────────────────────

def _val(v):
    if v is None or (isinstance(v, float) and util.isNan(v)):
        return ""
    return v


def _greek(greeks, field):
    if greeks is None:
        return ""
    val = getattr(greeks, field, None)
    if val is None or util.isNan(val):
        return ""
    return round(val, 4)


def read_option_data(ticker_cache):
    """Read current data from active streaming subscriptions."""
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


# ── Chain sheet ───────────────────────────────────────────────────

CALL_HEADERS = ["C Delta", "C Gamma", "C Theta", "C Volume", "C OI", "C Bid", "C Ask", "C Last"]
STRIKE_HEADER = ["Strike"]
PUT_HEADERS = ["P Last", "P Bid", "P Ask", "P OI", "P Volume", "P Delta", "P Gamma", "P Theta"]
ALL_HEADERS = CALL_HEADERS + STRIKE_HEADER + PUT_HEADERS


_chain_headers_written = False
_last_chain_data = None


def setup_excel(ws):
    """Clear data on ticker/expiry change, preserve row 3 (expiry selector).
    Headers + bold set once, not on every call."""
    global _chain_headers_written
    ws.range("A2:Q2").clear_contents()
    ws.range("A4:Q500").clear_contents()
    ws.range("A4").value = ALL_HEADERS
    if not _chain_headers_written:
        ws.range("A4:Q4").font.bold = True
        _chain_headers_written = True


def _reset_chain_cache():
    global _last_chain_data, _chain_headers_written
    _last_chain_data = None
    _chain_headers_written = False


def _build_chain_row(strike, call, put):
    """Build one row for the chain sheet."""
    return [
        call.get("delta", ""), call.get("gamma", ""), call.get("theta", ""),
        call.get("volume", ""), call.get("openInterest", ""),
        call.get("bid", ""), call.get("ask", ""), call.get("last", ""),
        strike,
        put.get("last", ""), put.get("bid", ""), put.get("ask", ""),
        put.get("openInterest", ""), put.get("volume", ""),
        put.get("delta", ""), put.get("gamma", ""), put.get("theta", ""),
    ]


def write_to_excel(ws, calls, puts, strikes, stock_price, expiry_display, prev_row_count):
    """Write option chain data to Excel. Skips rewrite if data unchanged."""
    global _last_chain_data

    filtered = _get_display_strikes(calls, puts, strikes, stock_price)

    data_rows = []
    for strike in filtered:
        data_rows.append(_build_chain_row(strike, calls.get(strike, {}), puts.get(strike, {})))

    # Skip rewrite if data hasn't changed
    cache_key = (round(stock_price, 2), expiry_display, tuple(tuple(r) for r in data_rows))
    if cache_key == _last_chain_data:
        return len(data_rows)
    _last_chain_data = cache_key

    num_rows = len(data_rows)

    if prev_row_count > num_rows:
        ws.range(f"A{5 + num_rows}:Q{4 + prev_row_count}").clear_contents()

    ws.range("A2").value = stock_price
    ws.range("B2").value = expiry_display

    if data_rows:
        ws.range("A5").value = data_rows

    return num_rows


# ── 10x sheet ─────────────────────────────────────────────────────

def _get_mid(data):
    """Get mid price from bid/ask data."""
    bid, ask = data.get("bid", ""), data.get("ask", "")
    if isinstance(bid, (int, float)) and isinstance(ask, (int, float)):
        return (bid + ask) / 2
    if isinstance(ask, (int, float)):
        return ask
    if isinstance(bid, (int, float)):
        return bid
    return None


def _calc_nx(mid, strike, stock_price, is_call):
    """Calculate % move required for each Nx return.

    Call at expiry: value = max(0, stock - strike). For Nx return, stock must reach
    strike + N*mid. % move = (target - current) / current * 100.

    Put at expiry: value = max(0, strike - stock). For Nx return, stock must fall to
    strike - N*mid. % move = (current - target) / current * 100.

    Returns "" for impossible scenarios (stock price < 0 required).
    """
    nx = []
    for m in MULTIPLIERS:
        if not mid or mid <= 0 or stock_price <= 0:
            nx.append("")
            continue
        if is_call:
            target = strike + m * mid
            pct = (target - stock_price) / stock_price * 100
        else:
            target = strike - m * mid
            if target < 0:
                nx.append("")
                continue
            pct = (stock_price - target) / stock_price * 100
        nx.append(round(pct, 1))
    return nx


def _pct_to_color(val, lo, hi):
    """Blue (low%/best) → white → red (high%/worst) gradient. Returns (r, g, b) 0-255."""
    if hi <= lo:
        return (220, 220, 220)
    t = max(0.0, min(1.0, (val - lo) / (hi - lo)))
    if t < 0.5:
        r = t * 2
        return (int(66 + r * 189), int(133 + r * 122), int(244 + r * 11))
    r = (t - 0.5) * 2
    return (int(255 - r * 21), int(255 - r * 188), int(255 - r * 202))


def _apply_colors_batch(sheet_name, color_cells):
    """Apply background colors via a single AppleScript call (Mac-compatible).
    color_cells: list of (cell_ref, r, g, b) where r,g,b are 0-255."""
    if not color_cells:
        return
    lines = []
    for ref, r, g, b in color_cells:
        lines.append(
            f'set color of interior object of range "{ref}" '
            f'to {{{r * 257}, {g * 257}, {b * 257}}}'
        )
    joined = "\n            ".join(lines)
    script = f'''tell application "Microsoft Excel"
    tell active workbook
        tell sheet "{sheet_name}"
            {joined}
        end tell
    end tell
end tell'''
    try:
        subprocess.run(["osascript", "-e", script], capture_output=True, timeout=30)
    except subprocess.TimeoutExpired:
        pass


# Layout: side-by-side like the chain sheet
# C 20x | ... | C 2x | C Delta | C Spread | C Ask | C Bid | Strike | P Bid | P Ask | P Spread | P Delta | P 2x | ... | P 20x

_NX_HDRS_REV = [f"C {m}x %" for m in reversed(MULTIPLIERS)]  # 20x..2x
_NX_HDRS_FWD = [f"P {m}x %" for m in MULTIPLIERS]             # 2x..20x
_10X_HEADERS = (_NX_HDRS_REV + ["C Delta", "C Spread", "C Ask", "C Bid"]
                + ["Strike"]
                + ["P Bid", "P Ask", "P Spread", "P Delta"] + _NX_HDRS_FWD)
_NUM_COLS = len(_10X_HEADERS)  # 23 columns, A through W
_LAST_COL = chr(64 + _NUM_COLS)  # 'W'

# Column indices for Nx% data (for color gradient)
_CALL_NX_COLS = list(range(len(MULTIPLIERS)))                                    # A-G
_PUT_NX_COLS = list(range(len(MULTIPLIERS) + 4 + 1 + 4, _NUM_COLS))             # Q-W


def _build_10x_row(strike, call_data, put_data, stock_price):
    """Build one side-by-side row for the 10x sheet.
    Call Nx% only shown for strikes >= current price (OTM/ATM calls).
    Put Nx% only shown for strikes <= current price (OTM/ATM puts)."""
    # Call side
    c_bid = call_data.get("bid", "")
    c_ask = call_data.get("ask", "")
    c_delta = call_data.get("delta", "")
    c_spread = round(c_ask - c_bid, 4) if isinstance(c_bid, (int, float)) and isinstance(c_ask, (int, float)) else ""
    c_mid = _get_mid(call_data)
    c_nx = _calc_nx(c_mid, strike, stock_price, True) if strike >= stock_price else [""] * len(MULTIPLIERS)

    # Put side
    p_bid = put_data.get("bid", "")
    p_ask = put_data.get("ask", "")
    p_delta = put_data.get("delta", "")
    p_spread = round(p_ask - p_bid, 4) if isinstance(p_bid, (int, float)) and isinstance(p_ask, (int, float)) else ""
    p_mid = _get_mid(put_data)
    p_nx = _calc_nx(p_mid, strike, stock_price, False) if strike <= stock_price else [""] * len(MULTIPLIERS)

    return (list(reversed(c_nx)) + [c_delta, c_spread, c_ask, c_bid]
            + [strike]
            + [p_bid, p_ask, p_spread, p_delta] + p_nx)


_last_10x_data = None
_10x_headers_written = False


def _reset_10x_cache():
    global _last_10x_data, _10x_headers_written
    _last_10x_data = None
    _10x_headers_written = False


def _setup_10x_headers(ws):
    """Write 10x headers + bold once. Idempotent — skips if already done."""
    global _10x_headers_written
    if _10x_headers_written:
        return
    ws.range("A2").value = [_10X_HEADERS]
    ws.range(f"A2:{_LAST_COL}2").font.bold = True
    _10x_headers_written = True


def write_10x_sheet(ws, calls, puts, stock_price, expiry_display):
    """Write the 10x multiplier analysis sheet. Only rewrites when data changes."""
    global _last_10x_data

    if not stock_price:
        return

    all_strikes = sorted(set(list(calls.keys()) + list(puts.keys())))
    if not all_strikes:
        return

    filtered = _get_display_strikes(calls, puts, all_strikes, stock_price)
    if not filtered:
        return

    # Build data rows
    data_rows = []
    for strike in filtered:
        data_rows.append(_build_10x_row(strike, calls.get(strike, {}), puts.get(strike, {}), stock_price))

    # Skip rewrite if data hasn't changed
    cache_key = (round(stock_price, 2), expiry_display, tuple(tuple(r) for r in data_rows))
    if cache_key == _last_10x_data:
        return
    _last_10x_data = cache_key

    # Clear old data + stale colors (clear() removes both values and formatting)
    data_start = 3
    end_row = data_start + max(len(data_rows), 50)
    ws.range(f"A{data_start}:{_LAST_COL}{end_row}").clear()

    # Row 1: info (values only — no formatting calls)
    ws.range("A1").value = f"Stock: ${stock_price:.2f}"
    ws.range("L1").value = f"Expiry: {expiry_display}"

    # Row 2: headers (written once, survives data clears since clear targets row 3+)
    _setup_10x_headers(ws)

    # Row 3+: data
    if data_rows:
        ws.range(f"A{data_start}").value = data_rows

    # ── Color gradient per Nx column ──
    color_cells = []
    for ci in _CALL_NX_COLS + _PUT_NX_COLS:
        col = chr(65 + ci)
        items = []
        for i, rd in enumerate(data_rows):
            if isinstance(rd[ci], (int, float)):
                items.append((data_start + i, rd[ci]))
        if len(items) < 2:
            continue

        lo = min(v for _, v in items)
        hi = max(v for _, v in items)
        for excel_row, v in items:
            r, g, b = _pct_to_color(v, lo, hi)
            color_cells.append((f"{col}{excel_row}", r, g, b))

    _apply_colors_batch("10x", color_cells)


# ── 10xtest sheet ─────────────────────────────────────────────────
# Same data as 10x, but highlights the minimum |value| in yellow
# for columns B-G (call Nx: 15x..2x) and S-W (put Nx: 5x..20x).

# Column indices for min-abs highlighting (0-based into the row list)
_TEST_MIN_COLS = [1, 2, 3, 4, 5, 6, 18, 19, 20, 21, 22]  # B-G, S-W
_YELLOW = (255, 255, 0)

_last_10xtest_data = None
_10xtest_headers_written = False


def _reset_10xtest_cache():
    global _last_10xtest_data, _10xtest_headers_written
    _last_10xtest_data = None
    _10xtest_headers_written = False


def write_10xtest_sheet(ws, calls, puts, stock_price, expiry_display):
    """Write 10x data with yellow highlight on the min absolute value per Nx column."""
    global _last_10xtest_data, _10xtest_headers_written

    if not stock_price:
        return

    all_strikes = sorted(set(list(calls.keys()) + list(puts.keys())))
    if not all_strikes:
        return

    filtered = _get_display_strikes(calls, puts, all_strikes, stock_price)
    if not filtered:
        return

    data_rows = []
    for strike in filtered:
        data_rows.append(_build_10x_row(strike, calls.get(strike, {}), puts.get(strike, {}), stock_price))

    cache_key = (round(stock_price, 2), expiry_display, tuple(tuple(r) for r in data_rows))
    if cache_key == _last_10xtest_data:
        return
    _last_10xtest_data = cache_key

    data_start = 3
    end_row = data_start + max(len(data_rows), 50)
    ws.range(f"A{data_start}:{_LAST_COL}{end_row}").clear()

    ws.range("A1").value = f"Stock: ${stock_price:.2f}"
    ws.range("L1").value = f"Expiry: {expiry_display}"

    if not _10xtest_headers_written:
        ws.range("A2").value = [_10X_HEADERS]
        ws.range(f"A2:{_LAST_COL}2").font.bold = True
        _10xtest_headers_written = True

    if data_rows:
        ws.range(f"A{data_start}").value = data_rows

    # Highlight min |value| per column in yellow
    color_cells = []
    for ci in _TEST_MIN_COLS:
        if ci >= _NUM_COLS:
            continue
        col = chr(65 + ci)
        items = []
        for i, rd in enumerate(data_rows):
            if isinstance(rd[ci], (int, float)):
                items.append((data_start + i, abs(rd[ci])))
        if not items:
            continue

        min_abs = min(v for _, v in items)
        for excel_row, abs_val in items:
            if abs_val == min_abs:
                color_cells.append((f"{col}{excel_row}", *_YELLOW))
                break  # only highlight the first occurrence

    _apply_colors_batch("10xtest", color_cells)


# ── Main ──────────────────────────────────────────────────────────

def _reset_state():
    """Reset all caches. Call on ticker/expiry change."""
    _reset_chain_cache()
    _reset_10x_cache()
    _reset_10xtest_cache()


def main():
    print("=" * 60)
    print("  TWS Option Chain Streamer (READ-ONLY)")
    print("  No trading activity — market data only")
    print("=" * 60)

    ib = connect_tws()

    wb = open_workbook()
    ws = wb.sheets["chain"]
    ws_10x = wb.sheets["10x"]
    ws_10xtest = wb.sheets["10xtest"]

    print("Ready. Type a ticker symbol in cell A1 of the 'chain' sheet.")
    print(f"Data refreshes every {REFRESH_SECONDS} seconds. Press Ctrl+C to stop.\n")

    last_ticker = None
    last_expiry = None
    ticker_cache = {}
    stock_ticker = None
    prev_row_count = 0
    stock_price = None
    all_strikes = None
    trading_class = None
    expiries = None

    def _full_reset():
        """Reset all loop state for reconnection or error recovery."""
        nonlocal last_ticker, last_expiry, ticker_cache, stock_ticker
        nonlocal prev_row_count, stock_price, all_strikes, trading_class, expiries
        last_ticker = None
        last_expiry = None
        ticker_cache = {}
        stock_ticker = None
        prev_row_count = 0
        stock_price = None
        all_strikes = None
        trading_class = None
        expiries = None
        _reset_state()

    try:
        while True:
            if not ib.isConnected():
                print("Connection lost. Reconnecting in 5 seconds...")
                time.sleep(5)
                try:
                    ib.disconnect()
                    ib = connect_tws()
                    _full_reset()
                except Exception as e:
                    print(f"  Reconnect failed: {e}")
                    time.sleep(REFRESH_SECONDS)
                    continue

            ticker_val = ws.range("A1").value
            if ticker_val is None or str(ticker_val).strip() == "":
                time.sleep(REFRESH_SECONDS)
                continue

            symbol = re.sub(r"[^A-Za-z]", "", str(ticker_val)).upper()
            if not symbol:
                time.sleep(REFRESH_SECONDS)
                continue

            try:
                # ── Ticker changed: qualify stock, get expiries ──
                if symbol != last_ticker:
                    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] New ticker: {symbol}")
                    _reset_state()

                    if ticker_cache:
                        cancel_subscriptions(ib, ticker_cache)

                    last_ticker = symbol
                    last_expiry = None

                    if stock_ticker is not None:
                        try:
                            ib.cancelMktData(stock_ticker.contract)
                        except Exception:
                            pass
                        stock_ticker = None

                    stock = Stock(symbol, "SMART", "USD")
                    _qualify_quietly(ib, stock)

                    if not stock.conId:
                        print(f"  ERROR: Could not find stock '{symbol}'")
                        ws.range("B2").value = f"ERROR: Unknown ticker '{symbol}'"
                        time.sleep(REFRESH_SECONDS)
                        last_ticker = None
                        continue

                    stock_ticker = subscribe_stock(ib, stock)
                    stock_price = get_stock_price(stock_ticker)
                    if stock_price is None:
                        print(f"  ERROR: Could not get price for '{symbol}'")
                        ws.range("B2").value = f"ERROR: No price for '{symbol}'"
                        time.sleep(REFRESH_SECONDS)
                        last_ticker = None
                        continue
                    print(f"  Current price: ${stock_price:.2f}")

                    expiries, all_strikes, trading_class = get_expiries(ib, stock, count=5)
                    if not expiries:
                        print(f"  ERROR: No options found for '{symbol}'")
                        ws.range("B2").value = f"ERROR: No options for '{symbol}'"
                        time.sleep(REFRESH_SECONDS)
                        last_ticker = None
                        continue

                    setup_expiry_selector(ws, expiries)
                    print(f"  Expiries: {', '.join(fmt_expiry(e) for e in expiries)}")

                # ── Read selected expiry (numeric 1-5 from B3) ──
                choice = _read_expiry_choice(ws)
                if not choice or not expiries or choice > len(expiries):
                    time.sleep(REFRESH_SECONDS)
                    continue

                expiry_api = expiries[choice - 1]
                expiry_display = fmt_expiry(expiry_api)

                if not all_strikes:
                    time.sleep(REFRESH_SECONDS)
                    continue

                # ── Expiry changed: re-subscribe to new chain ──
                if expiry_api != last_expiry:
                    print(f"  Switching to expiry: {expiry_display}")
                    _reset_state()

                    if ticker_cache:
                        cancel_subscriptions(ib, ticker_cache)

                    filtered_strikes = filter_strikes(all_strikes, stock_price)
                    print(f"  {expiry_display} | {len(filtered_strikes)} strikes (of {len(all_strikes)} total)")

                    subscribe_option_chain(ib, symbol, expiry_api, filtered_strikes, ticker_cache, trading_class)

                    setup_excel(ws)
                    prev_row_count = 0
                    last_expiry = expiry_api

                # ── Refresh cycle ──
                ib.sleep(0.1)

                if stock_ticker is not None:
                    live_price = get_stock_price(stock_ticker)
                    if live_price is not None:
                        stock_price = live_price

                print(f"[{datetime.now().strftime('%H:%M:%S')}] {symbol} ${stock_price:.2f} ({expiry_display})...", end=" ")
                calls, puts = read_option_data(ticker_cache)

                active_strikes = sorted({c.strike for c, _ in ticker_cache.values()})

                # Write chain sheet
                prev_row_count = write_to_excel(ws, calls, puts, active_strikes, stock_price, expiry_display, prev_row_count)

                # Write 10x sheets
                write_10x_sheet(ws_10x, calls, puts, stock_price, expiry_display)
                write_10xtest_sheet(ws_10xtest, calls, puts, stock_price, expiry_display)

                print(f"Done — {len(active_strikes)} strikes.")

            except (ConnectionError, OSError) as e:
                print(f"\n  Connection error: {e}")
                _full_reset()
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
