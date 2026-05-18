"""
ETF Corporate Action Adjuster
Adjusts historical NAV and Shares Outstanding for stock splits.
- Scans ±WINDOW_DAYS around the CA effective date to find the actual break
- Applies adjustment strictly before the detected break date
- Checks NAV and Shares independently
- Handles #N/A values around the break: ffill used only for detection, raw NaNs preserved in output
"""

import pandas as pd
import numpy as np

# ── CONFIG ──────────────────────────────────────────────────────────────────
NAV_FILE    = "nav.csv"
SHARES_FILE = "shares.csv"
CA_FILE     = "corporate_actions.csv"

OUT_NAV     = "nav_adjusted.csv"
OUT_SHARES  = "shares_adjusted.csv"
OUT_SUMMARY = "adjustment_summary.csv"

WINDOW_DAYS   = 3      # calendar days either side of effective date to scan
TOLERANCE_PCT = 0.05   # 5% — if observed jump is within this of expected, skip
# ────────────────────────────────────────────────────────────────────────────


def load_data():
    na_vals = ["#N/A", "N/A", "NA", "n/a", "#NA", ""]
    nav    = pd.read_csv(NAV_FILE,    index_col=0, parse_dates=True, na_values=na_vals, keep_default_na=True)
    shares = pd.read_csv(SHARES_FILE, index_col=0, parse_dates=True, na_values=na_vals, keep_default_na=True)
    ca     = pd.read_csv(CA_FILE,     parse_dates=["Effective Date"])
    return nav, shares, ca


def detect_break(series: pd.Series, event_date: pd.Timestamp,
                 expected_step: float) -> tuple[pd.Timestamp, float]:
    """
    Scan consecutive day-on-day ratios within ±WINDOW_DAYS of event_date.
    Return (break_date, observed_step) where break_date is the FIRST date
    after the jump (i.e., adjust everything strictly before it).
    Falls back to event_date if no data in window.

    expected_step:
      NAV    → ratio        (e.g. 0.0333 for 1:30 reverse split — price drops)
      Shares → 1/ratio      (e.g. 30    for 1:30 reverse split — count drops)
    """
    lo = event_date - pd.Timedelta(days=WINDOW_DAYS)
    hi = event_date + pd.Timedelta(days=WINDOW_DAYS)
    # Use ffill within the window ONLY for step detection — raw NaNs stay in the output.
    # This lets us see through #N/A gaps to find the true price/share discontinuity.
    # Extend one row before the window to seed ffill (handles #N/A at window boundary)
    pre_window = series[series.index < lo].dropna()
    seed = pre_window.iloc[[-1]] if len(pre_window) > 0 else pd.Series(dtype=float)
    raw_window = series[(series.index >= lo) & (series.index <= hi)]
    window = pd.concat([seed, raw_window]).ffill().dropna()
    window = window[window.index >= lo]  # drop seed row after ffill

    if len(window) < 2:
        return event_date, float("nan")

    best_date  = event_date
    best_step  = float("nan")
    best_dist  = float("inf")

    for i in range(1, len(window)):
        prev = window.iloc[i - 1]
        curr = window.iloc[i]
        if prev == 0:
            continue
        step = curr / prev
        dist = abs(step - expected_step)
        if dist < best_dist:
            best_dist = dist
            best_step = step
            best_date = window.index[i]

    return best_date, best_step


def check_and_adjust(series: pd.Series, event_date: pd.Timestamp,
                     ratio: float, label: str, ticker: str) -> dict:
    """
    1. Compute expected single-step jump for an unadjusted series.
    2. Find the actual break date within the window.
    3. If the jump matches expected (within tolerance) → needs adjustment.
    4. If already ~1.0 (no jump) → already adjusted, skip.
    Returns result dict and adjusted series (or original if skipped).
    """
    # Expected day-on-day ratio at the split date for an UNADJUSTED series:
    #   NAV:    price moves inversely to the split → step = 1/ratio
    #           (reverse split 1:30, ratio=0.033 → price jumps ×30)
    #           (forward split 2:1, ratio=2.0    → price drops ×0.5)
    #   Shares: count moves with the split → step = ratio
    #           (reverse split 1:30, ratio=0.033 → shares drop to ×0.033)
    #           (forward split 2:1, ratio=2.0    → shares double)
    expected_step = (1.0 / ratio) if label == "NAV" else ratio

    break_date, observed_step = detect_break(series, event_date, expected_step)

    result = {
        "ticker"              : ticker,
        "series"              : label,
        "ca_effective_date"   : event_date.date(),
        "detected_break_date" : break_date.date(),
        "date_offset_days"    : (break_date - event_date).days,
        "ratio"               : ratio,
        "expected_step"       : round(expected_step, 6),
        "observed_step"       : round(observed_step, 6) if not np.isnan(observed_step) else None,
        "needs_adjustment"    : False,
        "action_taken"        : "SKIPPED",
        "reason"              : ""
    }

    if np.isnan(observed_step):
        result["reason"]       = "Insufficient data in window to detect break"
        result["action_taken"] = "SKIPPED – no data in window"
        return result, series

    dev_from_expected = abs(observed_step - expected_step) / abs(expected_step)
    dev_from_one      = abs(observed_step - 1.0)

    if dev_from_one <= TOLERANCE_PCT:
        # No jump visible at the break → series already adjusted
        result["reason"]       = (f"No split jump detected: step={observed_step:.4f} ≈ 1.0 "
                                  f"(dev from 1: {dev_from_one:.2%}) — already adjusted")
        result["action_taken"] = "NO ACTION – already adjusted"
        return result, series

    elif dev_from_expected <= TOLERANCE_PCT:
        # Jump matches expected unadjusted discontinuity → needs adjustment
        result["needs_adjustment"] = True
        result["reason"]       = (f"Break detected: step={observed_step:.4f} ≈ "
                                  f"expected {expected_step:.4f} (dev {dev_from_expected:.2%})")
        result["action_taken"] = "ADJUSTED"
        factor = 1.0 / ratio if label == "NAV" else ratio
        adj = series.copy()
        # NaN rows multiply as NaN × factor = NaN, so they are preserved unchanged
        adj[adj.index < break_date] *= factor
        return result, adj

    else:
        # Jump exists but doesn't clearly match expected ratio — flag for review
        result["reason"]       = (f"Ambiguous: step={observed_step:.4f}, "
                                  f"expected {expected_step:.4f} (dev {dev_from_expected:.2%}), "
                                  f"dev from 1.0: {dev_from_one:.2%} — MANUAL REVIEW NEEDED")
        result["action_taken"] = "SKIPPED – ambiguous, review needed"
        return result, series


def process(df: pd.DataFrame, ca: pd.DataFrame, label: str):
    df_adj = df.copy().astype(float)
    summary = []

    for _, row in ca.iterrows():
        ticker     = row["Ticker"]
        event_date = row["Effective Date"]
        ratio      = float(row["Ratio"])

        if ticker not in df_adj.columns:
            summary.append({
                "ticker": ticker, "series": label,
                "ca_effective_date": event_date.date(),
                "detected_break_date": None, "date_offset_days": None,
                "ratio": ratio, "expected_step": None, "observed_step": None,
                "needs_adjustment": False,
                "action_taken": "SKIPPED – ticker missing",
                "reason": "Ticker not in dataset"
            })
            continue

        result, df_adj[ticker] = check_and_adjust(
            df_adj[ticker], event_date, ratio, label, ticker)
        summary.append(result)

        offset = f"offset={result['date_offset_days']:+d}d" if result['date_offset_days'] is not None else ""
        tag = f"[{result['action_taken'].split()[0]}]".ljust(12)
        print(f"  {tag} {ticker} | {label} | CA={event_date.date()} "
              f"break={result['detected_break_date']} {offset}")

    return df_adj, summary


def main():
    print("Loading data...")
    nav, shares, ca = load_data()
    print(f"NAV: {nav.shape}, Shares: {shares.shape}, CA rows: {len(ca)}\n")

    print("── Processing NAV ──────────────────────────────────────────────")
    nav_adj, nav_sum = process(nav, ca, "NAV")

    print("\n── Processing Shares ───────────────────────────────────────────")
    shr_adj, shr_sum = process(shares, ca, "Shares")

    nav_adj.round(6).to_csv(OUT_NAV)
    shr_adj.round(0).to_csv(OUT_SHARES)  # keep as float to preserve NaN rows

    cols = ["ticker","series","ca_effective_date","detected_break_date",
            "date_offset_days","ratio","expected_step","observed_step",
            "needs_adjustment","action_taken","reason"]
    summary_df = (pd.DataFrame(nav_sum + shr_sum)[cols]
                  .sort_values(["ticker","series"]))
    summary_df.to_csv(OUT_SUMMARY, index=False)

    print(f"\n── Summary ─────────────────────────────────────────────────────")
    print(summary_df.to_string(index=False))
    print(f"\nOutputs: {OUT_NAV}, {OUT_SHARES}, {OUT_SUMMARY}")


if __name__ == "__main__":
    main()
