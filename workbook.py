"""
Shared workbook setup for option_chain.py and scanner.py.

Opens twschain.xlsx (or creates it) and ensures all required sheets exist
with their headers. Safe to call from multiple scripts — only creates
sheets/headers if they don't already exist.
"""

from datetime import datetime

import xlwings as xw

WORKBOOK = "twschain.xlsx"


def fmt_expiry(exp):
    """Convert YYYYMMDD to human-readable '20-Feb-26' format."""
    return datetime.strptime(exp, "%Y%m%d").strftime("%d-%b-%y")


def open_workbook():
    """Open or create twschain.xlsx and ensure all sheets exist.
    Returns the workbook object."""
    try:
        wb = xw.Book(WORKBOOK)
    except Exception:
        wb = xw.Book()
        wb.save(WORKBOOK)

    existing = [s.name for s in wb.sheets]

    # Create required sheets
    for name in ["chain", "10x", "10xtest", "interesting", "signals_log"]:
        if name not in existing:
            wb.sheets.add(name, after=wb.sheets[-1])

    # Set up "interesting" headers
    ws_int = wb.sheets["interesting"]
    if ws_int.range("A1").value is None:
        ws_int.range("A1").value = "Tickers"
        ws_int.range("D1").value = "Surge %"
        ws_int.range("E1").value = 200
        ws_int.range("D2").value = "Drop %"
        ws_int.range("E2").value = 50
        ws_int.range("G1").value = "Last Scan"
        ws_int.range("J1").value = [[
            "Ticker", "Expiry", "Strike", "Type", "Current",
            "Peak", "Baseline", "Surge%", "Drop%", "Signal Time"
        ]]
        ws_int.range("J1:S1").font.bold = True

    # Set up "signals_log" headers
    ws_log = wb.sheets["signals_log"]
    if ws_log.range("A1").value is None:
        ws_log.range("A1").value = [[
            "Signal Time", "Ticker", "Expiry", "Strike", "Type",
            "Entry", "Peak", "Current", "PnL%", "Status"
        ]]
        ws_log.range("A1:J1").font.bold = True

    # Remove empty default sheets (Sheet1, etc.)
    if len(wb.sheets) > 5:
        for s in list(wb.sheets):
            if s.name.startswith("Sheet") and s.range("A1").value is None:
                try:
                    s.delete()
                except Exception:
                    pass

    wb.save()
    return wb
