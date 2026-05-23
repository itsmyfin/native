"""
esg_data.py — Data ingestion & analytics for ESG vs Plain ETF research note
============================================================================

Reads three Excel files and produces a clean etf_data dict per benchmark
pair that plugs directly into esg_report_pdf.build_pdf().

INPUT FILE FORMATS
------------------
1. INDEX PRICES  (e.g. index_prices.xlsx)
   - Sheet:   any name (default "Sheet1")
   - Col A:   Date  (any parseable date format)
   - Col B+:  Index ticker as header, daily closing price in USD

2. ETF NAV / TR  (e.g. etf_nav.xlsx)
   - Sheet:   any name (default "Sheet1")
   - Col A:   Date
   - Col B+:  ETF ticker as header, daily total-return NAV in USD

3. ETF FLOWS  (e.g. etf_flows.xlsx)
   - Col A:   ETF ticker (rows are tickers)
   - Row 1:   Date headers across columns
   - Body:    Flow values in USD millions

4. STATIC / METADATA  (second sheet in etf_nav.xlsx, or separate file)
   Columns (any order, case-insensitive):
       ticker | benchmark | listing_date | segment | esg
   'esg' column: "yes" / "no"  (case-insensitive)

OUTPUTS
-------
A dict keyed by benchmark index ticker, each value is an etf_data dict
ready for esg_report_pdf.build_pdf().

USAGE
-----
    from esg_data import load_all

    all_data = load_all(
        index_file  = "data/index_prices.xlsx",
        nav_file    = "data/etf_nav.xlsx",
        flows_file  = "data/etf_flows.xlsx",
        static_sheet= "Metadata",          # sheet name inside nav_file
        index_sheet = "Sheet1",
        nav_sheet   = "Sheet1",
    )

    for benchmark, etf_data in all_data.items():
        esg_report_pdf.build_pdf(etf_data, f"output/esg_note_{benchmark}.pdf")
"""

import warnings
import re
import numpy as np
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ─────────────────────────────────────────────────────────────────────────────
# 0. TICKER NORMALISATION
# ─────────────────────────────────────────────────────────────────────────────

# Bloomberg asset-class suffixes to strip, in order of specificity.
# Covers equities, fixed income, commodities, FX, indices across all regions.
_BBG_SUFFIXES = re.compile(
    r"\s+"                          # one or more spaces before suffix
    r"(?:"
    r"Equity|Comdty|Corp|Govt|Mtge|Muni|Pfd|"   # asset class
    r"Index|Curncy|"                              # index / FX
    r"US|LN|GY|FP|JP|HK|AU|CN|IN|"              # exchange country codes
    r"SM|IM|BB|NA|ID|SP|MM|VX|SQ|DC|NO|SS|FH|PW|GA|AV"  # more exchange codes
    r")\s*$",
    re.IGNORECASE,
)

def normalise_ticker(raw: str) -> str:
    """
    Strip Bloomberg exchange/asset-class suffixes from a ticker string.

    Examples
    --------
    'ESGU US Equity'  → 'ESGU'
    'ESGU US'         → 'ESGU'
    'ISF LN Equity'   → 'ISF'
    'SPX Index'       → 'SPX'
    'ESGU'            → 'ESGU'   (already clean — no change)
    'EUR Curncy'      → 'EUR'
    """
    ticker = str(raw).strip()
    # Iteratively strip suffixes until nothing changes
    # (handles multi-part suffixes like "US Equity")
    while True:
        cleaned = _BBG_SUFFIXES.sub("", ticker).strip()
        if cleaned == ticker:
            break
        ticker = cleaned
    return ticker.upper()


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Apply normalise_ticker to all column names of a DataFrame."""
    df.columns = [normalise_ticker(c) for c in df.columns]
    return df


def normalise_series_index(s: pd.Series) -> pd.Series:
    """Apply normalise_ticker to the index of a Series (used for flows rows)."""
    s.index = [normalise_ticker(i) for i in s.index]
    return s


# ─────────────────────────────────────────────────────────────────────────────
# 1. FILE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_prices(path: str, sheet: str = "Sheet1") -> pd.DataFrame:
    """
    Load index prices or ETF NAV.
    Returns DataFrame indexed by date, columns = normalised tickers.
    """
    df = pd.read_excel(path, sheet_name=sheet, index_col=0, parse_dates=True)
    df.index = pd.to_datetime(df.index)
    df = df.sort_index()
    df = normalise_columns(df)
    df = df.apply(pd.to_numeric, errors="coerce")
    return df


def load_static(path: str, sheet: str = "Metadata") -> pd.DataFrame:
    """
    Load static/metadata sheet.
    Returns DataFrame with columns: ticker, benchmark, listing_date,
    segment, esg (bool). Both ticker and benchmark are normalised
    (Bloomberg suffixes stripped) so they join cleanly with price data.
    """
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    # Normalise ticker and benchmark — strips "US Equity", "LN Equity" etc.
    df["ticker"]    = df["ticker"].apply(lambda x: normalise_ticker(str(x)))
    df["benchmark"] = df["benchmark"].apply(lambda x: normalise_ticker(str(x)))
    df["esg"]       = df["esg"].str.strip().str.lower() == "yes"
    if "listing_date" in df.columns:
        df["listing_date"] = pd.to_datetime(df["listing_date"], errors="coerce")
    return df.set_index("ticker")


def load_flows(path: str, sheet: str = "Sheet1") -> pd.DataFrame:
    """
    Load ETF flows.
    Expected format: Col A = Date, Col B+ = ETF ticker headers, rows = monthly flows.
    (Same orientation as the NAV and index price files.)
    Returns DataFrame indexed by date, columns = normalised tickers.
    """
    df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    df.index = pd.to_datetime(df.index, errors="coerce", dayfirst=False)
    df = df[df.index.notna()].sort_index()
    df = normalise_columns(df)          # strip Bloomberg suffixes from headers
    df = df.apply(pd.to_numeric, errors="coerce")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 2. RETURN CALCULATIONS
# ─────────────────────────────────────────────────────────────────────────────

def daily_returns(prices: pd.Series) -> pd.Series:
    """Simple daily return series from price/NAV series."""
    return prices.pct_change().dropna()


def annualised_return(prices: pd.Series, years: float) -> float:
    """
    Annualised total return over the most recent `years` period.
    Uses actual calendar days for precision.
    """
    end   = prices.last_valid_index()
    start = end - pd.DateOffset(years=int(years)) if years == int(years) \
            else end - pd.Timedelta(days=int(years * 365.25))
    # Find closest available date
    idx   = prices.index[prices.index >= start]
    if len(idx) == 0:
        return np.nan
    start_actual = idx[0]
    p_start = prices.loc[start_actual]
    p_end   = prices.loc[end]
    if p_start <= 0 or np.isnan(p_start):
        return np.nan
    n_years = (end - start_actual).days / 365.25
    return ((p_end / p_start) ** (1 / n_years) - 1) * 100


def annualised_return_table(esg_prices: pd.Series,
                            plain_prices: pd.Series,
                            horizons: list = [1, 3, 5]) -> dict:
    """
    Returns dict with keys 'labels', 'esg', 'plain' for the returns chart.
    Horizons in years. Skips horizons shorter than available history.
    """
    labels, esg_vals, plain_vals = [], [], []
    label_map = {1: "1 year", 2: "2 year", 3: "3 year", 5: "5 year", 10: "10 year"}

    # Align on common dates
    common = esg_prices.index.intersection(plain_prices.index)
    esg_c   = esg_prices.loc[common].dropna()
    plain_c = plain_prices.loc[common].dropna()

    for h in horizons:
        e = annualised_return(esg_c,   h)
        p = annualised_return(plain_c, h)
        if not (np.isnan(e) or np.isnan(p)):
            labels.append(label_map.get(h, f"{h}y"))
            esg_vals.append(round(e, 2))
            plain_vals.append(round(p, 2))

    return {"labels": labels, "esg": esg_vals, "plain": plain_vals}


def rolling_return_diff(esg_prices: pd.Series,
                        plain_prices: pd.Series,
                        window_months: int = 12) -> pd.Series:
    """
    Rolling annualised return of ESG minus plain.
    Resampled to month-end for cleaner charting.
    """
    common  = esg_prices.index.intersection(plain_prices.index)
    esg_m   = esg_prices.loc[common].resample("ME").last().dropna()
    plain_m = plain_prices.loc[common].resample("ME").last().dropna()

    window = window_months
    esg_roll   = esg_m.pct_change(window) * (12 / window) * 100
    plain_roll = plain_m.pct_change(window) * (12 / window) * 100
    diff = (esg_roll - plain_roll).dropna()
    return diff


def tracking_error(esg_prices: pd.Series,
                   plain_prices: pd.Series) -> float:
    """
    Annualised tracking error (std of daily return differences * sqrt(252)).
    Returns value in % (e.g. 1.23 = 1.23%).
    """
    common    = esg_prices.index.intersection(plain_prices.index)
    esg_ret   = daily_returns(esg_prices.loc[common].dropna())
    plain_ret = daily_returns(plain_prices.loc[common].dropna())
    diff      = esg_ret - plain_ret
    return round(diff.std() * np.sqrt(252) * 100, 3)


def rolling_tracking_error(esg_prices: pd.Series,
                            plain_prices: pd.Series,
                            window_months: int = 12) -> pd.Series:
    """
    Rolling annualised tracking error, resampled to month-end.
    """
    common    = esg_prices.index.intersection(plain_prices.index)
    esg_ret   = daily_returns(esg_prices.loc[common].dropna())
    plain_ret = daily_returns(plain_prices.loc[common].dropna())
    diff      = esg_ret - plain_ret
    window_days = window_months * 21   # ~21 trading days/month
    roll_te   = diff.rolling(window_days).std() * np.sqrt(252) * 100
    return roll_te.resample("ME").last().dropna()


def sharpe_ratio(prices: pd.Series, rf_annual: float = 0.04) -> float:
    """
    Annualised Sharpe ratio. rf_annual = risk-free rate (e.g. 0.04 = 4%).
    """
    ret = daily_returns(prices.dropna())
    rf_daily = rf_annual / 252
    excess   = ret - rf_daily
    if excess.std() == 0:
        return np.nan
    return round((excess.mean() / excess.std()) * np.sqrt(252), 3)


def max_drawdown(prices: pd.Series) -> float:
    """Maximum drawdown as a positive percentage (e.g. 18.4 = 18.4% drawdown)."""
    roll_max = prices.cummax()
    drawdown = (prices - roll_max) / roll_max
    return round(abs(drawdown.min()) * 100, 2)


def volatility_annual(prices: pd.Series) -> float:
    """Annualised volatility (%) from daily returns."""
    return round(daily_returns(prices.dropna()).std() * np.sqrt(252) * 100, 2)


def performance_summary_table(esg_prices: pd.Series,
                               plain_prices: pd.Series,
                               esg_ticker: str,
                               plain_ticker: str,
                               rf_annual: float = 0.04) -> pd.DataFrame:
    """
    Full performance summary table for both ETFs.
    Returns a DataFrame suitable for printing or PDF table.
    """
    rows = []
    for label, prices in [(esg_ticker, esg_prices), (plain_ticker, plain_prices)]:
        row = {
            "Ticker":     label,
            "1y Ann Ret (%)":  annualised_return(prices, 1),
            "3y Ann Ret (%)":  annualised_return(prices, 3),
            "5y Ann Ret (%)":  annualised_return(prices, 5),
            "Ann Vol (%)":     volatility_annual(prices),
            "Sharpe":          sharpe_ratio(prices, rf_annual),
            "Max DD (%)":      max_drawdown(prices),
        }
        rows.append(row)

    # Add a differential row
    diff_row = {"Ticker": "Differential (ESG - Plain)"}
    for col in ["1y Ann Ret (%)", "3y Ann Ret (%)", "5y Ann Ret (%)",
                "Ann Vol (%)", "Sharpe", "Max DD (%)"]:
        try:
            diff_row[col] = round(rows[0][col] - rows[1][col], 2)
        except Exception:
            diff_row[col] = np.nan
    rows.append(diff_row)

    df = pd.DataFrame(rows).set_index("Ticker")
    return df.round(2)


def flows_vs_diff(flows_series: pd.Series,
                  rolling_diff: pd.Series) -> dict:
    """
    Align monthly flows with rolling differential on common month-end dates.
    Returns dict with keys 'dates', 'net_flows', 'rolling_diff'.
    """
    flows_m = flows_series.resample("ME").sum()
    common  = flows_m.index.intersection(rolling_diff.index)
    if len(common) == 0:
        return {"dates": [], "net_flows": [], "rolling_diff": []}
    return {
        "dates":        common.tolist(),
        "net_flows":    flows_m.loc[common].tolist(),
        "rolling_diff": rolling_diff.loc[common].tolist(),
    }


def detect_regime_annotations(diff: pd.Series,
                               top_n: int = 2) -> list:
    """
    Auto-detect the most extreme peaks and troughs in the differential
    for annotation on the rolling diff chart.
    Returns list of (integer_index, label) tuples.
    """
    if len(diff) == 0:
        return []
    annotations = []
    peak_idx = int(diff.values.argmax())
    trough_idx = int(diff.values.argmin())
    annotations.append((peak_idx,   f"Peak +{diff.iloc[peak_idx]:.1f}%"))
    annotations.append((trough_idx, f"Trough {diff.iloc[trough_idx]:.1f}%"))
    return annotations


# ─────────────────────────────────────────────────────────────────────────────
# 3. PAIR BUILDER — one ESG + one plain per benchmark
# ─────────────────────────────────────────────────────────────────────────────

def best_esg_plain_pair(static: pd.DataFrame,
                         benchmark: str) -> tuple[str, str] | tuple[None, None]:
    """
    For a given benchmark, return (esg_ticker, plain_ticker).
    Picks the ESG ETF with the longest history and the plain ETF with
    the largest AUM (proxied by longest history if AUM not available).
    Returns (None, None) if no valid pair found.
    """
    subset = static[static["benchmark"] == benchmark]
    esg_tickers   = subset[subset["esg"]].index.tolist()
    plain_tickers = subset[~subset["esg"]].index.tolist()
    if not esg_tickers or not plain_tickers:
        return None, None
    # Prefer earliest listing date
    def earliest(tickers):
        if "listing_date" in subset.columns:
            dated = subset.loc[tickers, "listing_date"].dropna()
            if len(dated):
                return dated.idxmin()
        return tickers[0]
    return earliest(esg_tickers), earliest(plain_tickers)


# ─────────────────────────────────────────────────────────────────────────────
# 4. MAIN LOADER — returns dict of etf_data per benchmark
# ─────────────────────────────────────────────────────────────────────────────

def load_all(index_file: str,
             nav_file: str,
             flows_file: str,
             static_sheet: str = "Metadata",
             index_sheet: str  = "Sheet1",
             nav_sheet: str    = "Sheet1",
             flows_sheet: str  = "Sheet1",
             horizons: list    = [1, 3, 5],
             rf_annual: float  = 0.04,
             rolling_window: int = 12) -> dict:
    """
    Master loader. Returns dict keyed by benchmark ticker:
        { "SP500": etf_data_dict, "MXEF": etf_data_dict, ... }

    Each etf_data_dict is ready for esg_report_pdf.build_pdf().
    Also attaches extra analytics under key '_analytics' for reference.
    """
    print("Loading data files...")
    index_prices = load_prices(index_file,  index_sheet)
    etf_nav      = load_prices(nav_file,    nav_sheet)
    static       = load_static(nav_file,    static_sheet)
    flows        = load_flows(flows_file,   flows_sheet)

    print(f"  Index prices:  {index_prices.shape[1]} series, "
          f"{index_prices.index[0].date()} → {index_prices.index[-1].date()}")
    print(f"  ETF NAV:       {etf_nav.shape[1]} tickers")
    print(f"  Flows:         {flows.shape[1]} tickers")
    print(f"  Metadata:      {len(static)} ETFs across "
          f"{static['benchmark'].nunique()} benchmarks")

    benchmarks = static["benchmark"].unique()
    all_data   = {}

    for bm in benchmarks:
        esg_ticker, plain_ticker = best_esg_plain_pair(static, bm)
        if esg_ticker is None:
            print(f"  [{bm}] No valid ESG/plain pair — skipping")
            continue

        # ── Pull price series
        if esg_ticker not in etf_nav.columns:
            print(f"  [{bm}] ESG ticker {esg_ticker} not in NAV file — skipping")
            continue
        if plain_ticker not in etf_nav.columns:
            print(f"  [{bm}] Plain ticker {plain_ticker} not in NAV file — skipping")
            continue

        esg_px   = etf_nav[esg_ticker].dropna()
        plain_px = etf_nav[plain_ticker].dropna()

        # Trim to common history
        start = max(esg_px.index[0], plain_px.index[0])
        esg_px   = esg_px[esg_px.index >= start]
        plain_px = plain_px[plain_px.index >= start]

        print(f"\n  [{bm}] {esg_ticker} (ESG) vs {plain_ticker} (plain) | "
              f"from {start.date()}")

        # ── Returns table
        ret_table = annualised_return_table(esg_px, plain_px, horizons)

        # ── Rolling differential
        roll_diff = rolling_return_diff(esg_px, plain_px, rolling_window)
        annotations = detect_regime_annotations(roll_diff)

        # ── Flows
        esg_flows = flows[esg_ticker] if esg_ticker in flows.columns \
                    else pd.Series(dtype=float)
        flows_data = flows_vs_diff(esg_flows, roll_diff) if len(esg_flows) \
                     else {"dates": [], "net_flows": [], "rolling_diff": []}

        # ── Analytics (not needed for charts but useful for tables/notes)
        te       = tracking_error(esg_px, plain_px)
        roll_te  = rolling_tracking_error(esg_px, plain_px, rolling_window)
        perf_tbl = performance_summary_table(
            esg_px, plain_px, esg_ticker, plain_ticker, rf_annual)

        print(f"         Tracking error (full period): {te:.2f}%")
        print(f"         Performance summary:\n{perf_tbl.to_string()}")

        # ── Sector delta placeholder (requires holdings data not in scope)
        # Populate from your holdings source when available.
        sector_delta_placeholder = {
            "sectors": ["Technology", "Healthcare", "Utilities",
                        "Consumer Disc", "Industrials", "Financials",
                        "Materials", "Energy"],
            "deltas":  [0.0] * 8,   # ← replace with real holdings diff
        }

        # ── Drift placeholder (requires quarterly holdings snapshots)
        drift_placeholder = {
            "dates":        [start.year],
            "overlap_pct":  [np.nan],
            "esg_energy":   [np.nan],
            "plain_energy": [np.nan],
        }

        # ── Assemble etf_data dict
        etf_data = {
            "benchmark":    bm,
            "esg_ticker":   esg_ticker,
            "plain_ticker": plain_ticker,

            # Chart 1 — annualised returns
            "returns": ret_table,

            # Chart 2 — rolling differential
            "rolling_diff": {
                "dates":       roll_diff.index.tolist(),
                "values":      roll_diff.values.tolist(),
                "annotations": annotations,
            },

            # Chart 3 — sector delta (populate from holdings when available)
            "sector_delta": sector_delta_placeholder,

            # Chart 4 — mandate drift (populate from holdings when available)
            "drift": drift_placeholder,

            # Chart 5 — flows vs differential
            "flows": flows_data,

            # Extra analytics (for PDF tables, text auto-population)
            "_analytics": {
                "tracking_error_pct":   te,
                "rolling_te":           roll_te,
                "performance_summary":  perf_tbl,
                "esg_sharpe":           sharpe_ratio(esg_px, rf_annual),
                "plain_sharpe":         sharpe_ratio(plain_px, rf_annual),
                "esg_vol":              volatility_annual(esg_px),
                "plain_vol":            volatility_annual(plain_px),
                "esg_maxdd":            max_drawdown(esg_px),
                "plain_maxdd":          max_drawdown(plain_px),
                "history_start":        start,
            },
        }

        all_data[bm] = etf_data

    print(f"\nLoaded {len(all_data)} benchmark pairs.")
    return all_data


# ─────────────────────────────────────────────────────────────────────────────
# 5. DIAGNOSTIC — print a summary without generating PDFs
# ─────────────────────────────────────────────────────────────────────────────

def print_diagnostics(all_data: dict):
    """Print a quick stats summary for all benchmarks."""
    print("\n" + "=" * 70)
    print("DIAGNOSTIC SUMMARY")
    print("=" * 70)
    for bm, d in all_data.items():
        a = d["_analytics"]
        print(f"\n  {bm}: {d['esg_ticker']} vs {d['plain_ticker']}")
        print(f"    History from: {a['history_start'].date()}")
        print(f"    Tracking error: {a['tracking_error_pct']:.2f}%")
        print(f"    ESG  — Sharpe: {a['esg_sharpe']:.2f} | "
              f"Vol: {a['esg_vol']:.1f}% | MaxDD: {a['esg_maxdd']:.1f}%")
        print(f"    Plain— Sharpe: {a['plain_sharpe']:.2f} | "
              f"Vol: {a['plain_vol']:.1f}% | MaxDD: {a['plain_maxdd']:.1f}%")
        ret = d["returns"]
        for i, lbl in enumerate(ret["labels"]):
            print(f"    {lbl}: ESG {ret['esg'][i]:+.1f}% | "
                  f"Plain {ret['plain'][i]:+.1f}% | "
                  f"Diff {ret['esg'][i]-ret['plain'][i]:+.1f}%")
