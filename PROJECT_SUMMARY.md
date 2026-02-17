# TWS Option Chain Streamer — Project Summary

## What It Does
Python script that streams real-time option chain data from Interactive Brokers TWS into Excel. Type a ticker in cell A1 of `twschain.xlsx` (sheet: `chain`) and the nearest-expiry option chain populates automatically every 10 seconds.

## Excel Layout
- **A1**: Ticker symbol (user types this)
- **A2**: Live stock price (updates every 10s)
- **A3**: Expiry date
- **Row 4**: Headers
- **Row 5+**: Option chain data

### Columns (left to right)
Calls: Delta, Gamma, Theta, Volume, OI, Bid, Ask, Last | **Strike** | Puts: Last, Bid, Ask, OI, Volume, Delta, Gamma, Theta

## Architecture
```
TWS (port 7496) ← ib_insync (streaming, read-only) → Python → xlwings → Excel
```

- **Streaming subscriptions** — subscribed once per ticker, read every 10s. No snapshots, no per-request fees.
- **Read-only** — connects with `readonly=True`, no order/trading capability.
- **Minimal Excel interaction** — 2-3 AppleScript calls per refresh cycle to prevent freezing.

## Key Design Decisions
- `tradingClass` specified when building option contracts to avoid ambiguous matches (e.g., PLTR vs 2PLTR)
- Strikes filtered to +/-20% of stock price, capped at 80 strikes centered around ATM
- Total TWS market data lines capped at 100 (TWS account limit)
- Stock price kept as a separate streaming subscription for live updates
- Ticker input sanitized to ASCII letters only (Excel sometimes reads garbled Unicode)
- Auto-reconnect on TWS disconnection

## Files
- `option_chain.py` — main script
- `requirements.txt` — dependencies (`ib_insync`, `xlwings`)
- `.gitignore` — excludes `.claude/`, `__pycache__/`, `.env`

## Issues Encountered & Resolved
1. **Snapshot + generic ticks incompatible** — TWS error 321. Switched to streaming `reqMktData`.
2. **Invalid far-OTM strikes** — TWS error 200. Fixed with price-based strike filtering.
3. **Ambiguous contracts** — Stocks like PLTR have multiple option classes. Fixed by specifying `tradingClass` from `reqSecDefOptParams`.
4. **Excel freezing** — Too many AppleScript calls per cycle. Reduced to 2-3 bulk writes.
5. **`end("down")` crash** — Returns row 1,048,576 on empty sheets. Switched to `used_range`.
6. **"Unknown contract" spam** — Suppressed by redirecting stdout during `qualifyContracts`.
7. **TWS market data line limit** — Capped subscriptions at 100 total.

---

## To Do — Future Work

### High Priority
- [ ] **Expiry selector** — Allow picking a specific expiry instead of always using the nearest one (e.g., read from cell B1 or a dropdown)
- [ ] **Multiple expiries** — Show 2-3 nearest expiries on separate sheets or side by side
- [ ] **IV column** — Add implied volatility to the chain display
- [ ] **Highlight ATM strike** — Bold or color the row closest to current stock price

### Medium Priority
- [ ] **Auto-refresh strikes on large price move** — If stock moves >5% from the price used to filter strikes, re-subscribe with updated range
- [ ] **Watchlist mode** — Support multiple tickers, each on its own sheet
- [ ] **Bid/ask midpoint column** — Useful for quick fair value reference
- [ ] **OI change tracking** — Show delta in OI from previous session
- [ ] **Color coding** — ITM vs OTM background shading, bid/ask spread heat map

### Low Priority / Nice to Have
- [ ] **Logging to file** — Write debug output to a log file instead of console
- [ ] **Config file** — Move constants (port, refresh rate, max strikes) to a YAML/JSON config
- [ ] **Google Sheets alternative** — Port to Google Sheets using gspread for cloud access
- [ ] **Historical chain snapshots** — Periodically save chain state to CSV for later analysis
- [ ] **Earnings flag** — Highlight when the selected expiry spans an earnings date
