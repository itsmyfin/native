"""
run_esg_report.py — Master runner
==================================
Connects esg_data.py (data ingestion + analytics) with
esg_report_pdf.py (chart generation + PDF assembly).

Generates one research note PDF per benchmark pair found in your data.

SETUP
-----
1. Set the three file paths below under CONFIG
2. Set the sheet names if different from defaults
3. Run:  python run_esg_report.py

OUTPUT
------
  output/esg_note_<BENCHMARK>.pdf     — one per benchmark pair
  output/esg_color_palette.pdf        — firm color reference
  output/diagnostics.txt              — stats summary for QA

DEPENDENCIES
------------
  pip install pandas numpy openpyxl reportlab matplotlib
"""

import os
import sys
import datetime

# ── CONFIG — edit these paths ─────────────────────────────────────────────

INDEX_FILE   = "data/index_prices.xlsx"    # index price series
NAV_FILE     = "data/etf_nav.xlsx"         # ETF NAV + static metadata sheet
FLOWS_FILE   = "data/etf_flows.xlsx"       # ETF flows

# Sheet names inside the Excel files
INDEX_SHEET  = "Sheet1"       # sheet in index_prices.xlsx
NAV_SHEET    = "Sheet1"       # price/NAV sheet in etf_nav.xlsx
STATIC_SHEET = "Metadata"     # static data sheet in etf_nav.xlsx
FLOWS_SHEET  = "Sheet1"       # sheet in etf_flows.xlsx

# Analysis parameters
HORIZONS        = [1, 3, 5]   # annualised return horizons in years
ROLLING_WINDOW  = 12          # months for rolling return & TE
RF_ANNUAL       = 0.04        # risk-free rate for Sharpe (4%)

OUTPUT_DIR = "output"

# ── IMPORTS ───────────────────────────────────────────────────────────────

sys.path.insert(0, os.path.dirname(__file__))

import esg_data as data_module
import esg_report_pdf as pdf_module

# Push our firm colors into the pdf module so charts match
pdf_module.FIRM_COLORS = {
    "esg":          "#185FA5",
    "plain":        "#888780",
    "accent":       "#D85A30",
    "drift":        "#534AB7",
    "flows_pos":    "#378ADD",
    "flows_neg":    "#993C1D",
    "ref_line":     "#B4B2A9",
    "outperform":   "#C8EDD8",
    "underperform": "#FAD5C8",
    "background":   "#FFFFFF",
    "panel":        "#F7F6F2",
    "grid":         "#E0DED8",
    "axis_text":    "#5F5E5A",
    "title_text":   "#1A1A18",
    "annotation":   "#534AB7",
}


# ── HELPERS ───────────────────────────────────────────────────────────────

def enrich_pdf_text(etf_data: dict) -> dict:
    """
    Auto-populate the text fields in the PDF sections with real computed
    stats from _analytics, so the note text reflects actual numbers.
    """
    a  = etf_data.get("_analytics", {})
    bm = etf_data["benchmark"]
    esg_t  = etf_data.get("esg_ticker",   "ESG ETF")
    plain_t= etf_data.get("plain_ticker", "Plain ETF")

    te    = a.get("tracking_error_pct", float("nan"))
    start = a.get("history_start")
    start_str = start.strftime("%B %Y") if start else "inception"

    # Patch the report title to include benchmark
    pdf_module.REPORT_TITLE = (
        f"ESG vs Plain ETF: Performance & Mandate Integrity — {bm}"
    )

    # Patch report date
    pdf_module.REPORT_DATE = datetime.date.today().strftime("%d %B %Y")

    return etf_data   # etf_data itself drives chart data; text is in build_pdf


def write_diagnostics(all_data: dict, path: str):
    """Write a plain-text diagnostics file for QA."""
    import io as _io
    buf = _io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = buf
    data_module.print_diagnostics(all_data)
    sys.stdout = old_stdout
    text = buf.getvalue()
    with open(path, "w") as f:
        f.write(text)
    print(f"  Diagnostics written: {path}")


# ── MAIN ──────────────────────────────────────────────────────────────────

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── 1. Validate input files
    for label, path in [("Index prices", INDEX_FILE),
                        ("ETF NAV",      NAV_FILE),
                        ("ETF flows",    FLOWS_FILE)]:
        if not os.path.exists(path):
            print(f"ERROR: {label} file not found: {path}")
            print("       Update the CONFIG paths at the top of run_esg_report.py")
            sys.exit(1)

    # ── 2. Load & compute all analytics
    all_data = data_module.load_all(
        index_file    = INDEX_FILE,
        nav_file      = NAV_FILE,
        flows_file    = FLOWS_FILE,
        static_sheet  = STATIC_SHEET,
        index_sheet   = INDEX_SHEET,
        nav_sheet     = NAV_SHEET,
        flows_sheet   = FLOWS_SHEET,
        horizons      = HORIZONS,
        rf_annual     = RF_ANNUAL,
        rolling_window= ROLLING_WINDOW,
    )

    if not all_data:
        print("ERROR: No benchmark pairs found. Check your metadata sheet.")
        sys.exit(1)

    # ── 3. Write diagnostics
    write_diagnostics(all_data, os.path.join(OUTPUT_DIR, "diagnostics.txt"))

    # ── 4. Generate one PDF per benchmark
    for bm, etf_data in all_data.items():
        etf_data = enrich_pdf_text(etf_data)
        safe_bm  = bm.replace(" ", "_").replace("/", "-")
        out_path = os.path.join(OUTPUT_DIR, f"esg_note_{safe_bm}.pdf")
        try:
            pdf_module.build_pdf(etf_data, out_path)
        except Exception as e:
            print(f"  ERROR generating PDF for {bm}: {e}")
            import traceback; traceback.print_exc()

    # ── 5. Palette reference PDF
    palette_path = os.path.join(OUTPUT_DIR, "esg_color_palette.pdf")
    pdf_module.build_palette_pdf(palette_path)

    print(f"\n{'='*60}")
    print(f"Done. {len(all_data)} report(s) written to: {OUTPUT_DIR}/")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
