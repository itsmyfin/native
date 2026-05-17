"""
ETF Corporate Action Adjuster
Adjusts historical NAV and Shares Outstanding for stock splits.
Applies sanity checks independently for each series before adjusting.
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime

# ── CONFIG ──────────────────────────────────────────────────────────────────
NAV_FILE    = "nav.csv"
SHARES_FILE = "shares.csv"
CA_FILE     = "corporate_actions.csv"

OUT_NAV     = "nav_adjusted.csv"
OUT_SHARES  = "shares_adjusted.csv"
OUT_SUMMARY = "adjustment_summary.csv"

# Tolerance: if pre/post-event ratio is within this % of expected, skip adjust
TOLERANCE_PCT = 0.02   # 2%
# Minimum rows required on each side of the event date for sanity check
MIN_ROWS_EACH_SIDE = 1
# ────────────────────────────────────────────────────────────────────────────


def load_data():
    nav    = pd.read_csv(NAV_FILE,    index_col=0, parse_dates=True)
    shares = pd.read_csv(SHARES_FILE, index_col=0, parse_dates=True)
    ca     = pd.read_csv(CA_FILE,     parse_dates=["Effective Date"])
    nav.index    = pd.to_datetime(nav.index)
    shares.index = pd.to_datetime(shares.index)
    return nav, shares, ca


def sanity_check(series: pd.Series, event_date: pd.Timestamp, ratio: float,
                 label: str, ticker: str) -> dict:
    """
    Check whether the series already looks adjusted around event_date.
    
    For a SPLIT with ratio < 1 (reverse split, e.g. 1:30 → ratio=0.0333):
      - NAV   should JUMP up  after effective date (post/pre ~ 1/ratio)
      - Shares should DROP    after effective date (post/pre ~ ratio)
    For ratio > 1 (forward split):
      - NAV   should DROP     after effective date
      - Shares should JUMP up after effective date

    Returns dict with keys: needs_adjustment (bool), reason (str), details (dict)
    """
    pre  = series[series.index <  event_date].dropna()
    post = series[series.index >= event_date].dropna()

    result = {
        "ticker": ticker, "series": label, "effective_date": event_date.date(),
        "ratio": ratio, "needs_adjustment": False, "reason": "", "action_taken": "SKIPPED",
        "pre_mean": None, "post_mean": None, "observed_ratio": None, "expected_ratio": None
    }

    if len(pre) < MIN_ROWS_EACH_SIDE or len(post) < MIN_ROWS_EACH_SIDE:
        result["reason"] = f"Insufficient data ({len(pre)} pre, {len(post)} post rows)"
        result["action_taken"] = "SKIPPED – insufficient data"
        return result

    pre_mean  = pre.iloc[-min(5, len(pre)):].mean()   # last 5 pre rows
    post_mean = post.iloc[:min(5, len(post))].mean()  # first 5 post rows

    result["pre_mean"]  = round(pre_mean,  6)
    result["post_mean"] = round(post_mean, 6)

    if pre_mean == 0:
        result["reason"] = "Pre-event mean is zero; cannot compute ratio"
        result["action_taken"] = "SKIPPED – zero pre-mean"
        return result

    observed = post_mean / pre_mean

    # Split-adjusted = pre-event history brought DOWN to match post-event level.
    # An unadjusted series shows a visible jump at the event date:
    #   NAV:    post/pre ≈ ratio      (price drops by the split factor)
    #   Shares: post/pre ≈ 1/ratio   (share count drops post-split)
    # Already-adjusted series has post/pre ≈ 1.0 (no jump visible).
    # We flag for adjustment when observed matches the unadjusted expected ratio.
    if label == "NAV":
        expected = ratio
    else:  # Shares
        expected = 1.0 / ratio

    result["observed_ratio"] = round(observed, 6)
    result["expected_ratio"] = round(expected, 6)

    deviation = abs(observed - expected) / expected

    if deviation <= TOLERANCE_PCT:
        result["reason"] = (f"Already adjusted: observed ratio {observed:.4f} ≈ "
                            f"expected {expected:.4f} (dev {deviation:.2%})")
        result["action_taken"] = "NO ACTION – already adjusted"
    else:
        # Also check if series looks unadjusted (ratio ≈ 1, i.e. no jump at all)
        deviation_from_one = abs(observed - 1.0)
        result["needs_adjustment"] = True
        result["reason"] = (f"Adjustment needed: observed ratio {observed:.4f}, "
                            f"expected {expected:.4f} (dev {deviation:.2%})")
        result["action_taken"] = "ADJUSTED"

    return result


def apply_adjustment(series: pd.Series, event_date: pd.Timestamp,
                     ratio: float, label: str) -> pd.Series:
    """Multiply all pre-event values by the back-adjustment factor."""
    s = series.copy()
    mask = s.index < event_date
    if label == "NAV":
        factor = 1.0 / ratio   # divide old NAV by ratio → brings pre-split prices DOWN
    else:
        factor = ratio         # multiply old shares by ratio → brings pre-split count UP
    s[mask] = s[mask] * factor
    return s


def process(df: pd.DataFrame, ca: pd.DataFrame, label: str):
    """
    Process one dataframe (NAV or Shares) against all corporate actions.
    Returns adjusted df and list of summary dicts.
    """
    df_adj = df.copy().astype(float)
    summary_rows = []

    for _, row in ca.iterrows():
        ticker      = row["Ticker"]
        event_date  = row["Effective Date"]
        ratio       = float(row["Ratio"])

        if ticker not in df_adj.columns:
            summary_rows.append({
                "ticker": ticker, "series": label,
                "effective_date": event_date.date(), "ratio": ratio,
                "needs_adjustment": False, "reason": "Ticker not in dataset",
                "action_taken": "SKIPPED – ticker missing",
                "pre_mean": None, "post_mean": None,
                "observed_ratio": None, "expected_ratio": None
            })
            continue

        check = sanity_check(df_adj[ticker], event_date, ratio, label, ticker)
        summary_rows.append(check)

        if check["needs_adjustment"]:
            df_adj[ticker] = apply_adjustment(df_adj[ticker], event_date, ratio, label)
            print(f"  [ADJUSTED]  {ticker} | {label} | {event_date.date()} | ratio={ratio}")
        else:
            print(f"  [SKIPPED]   {ticker} | {label} | {event_date.date()} | {check['reason']}")

    return df_adj, summary_rows


def main():
    print("Loading data...")
    nav, shares, ca = load_data()

    print(f"\nNAV shape: {nav.shape}, Shares shape: {shares.shape}, CA rows: {len(ca)}\n")
    print(f"Corporate actions:\n{ca.to_string(index=False)}\n")

    print("── Processing NAV ──────────────────────────────────────────────")
    nav_adj, nav_summary = process(nav, ca, "NAV")

    print("\n── Processing Shares ───────────────────────────────────────────")
    shares_adj, shares_summary = process(shares, ca, "Shares")

    # Round NAV to 6dp, shares to 0dp
    nav_adj    = nav_adj.round(6)
    shares_adj = shares_adj.round(0).astype(int)

    nav_adj.to_csv(OUT_NAV)
    shares_adj.to_csv(OUT_SHARES)

    summary_df = pd.DataFrame(nav_summary + shares_summary)
    col_order  = ["ticker","series","effective_date","ratio","needs_adjustment",
                  "action_taken","pre_mean","post_mean","observed_ratio","expected_ratio","reason"]
    summary_df = summary_df[col_order].sort_values(["ticker","series"])
    summary_df.to_csv(OUT_SUMMARY, index=False)

    print(f"\n── Summary ─────────────────────────────────────────────────────")
    print(summary_df.to_string(index=False))
    print(f"\nOutputs written: {OUT_NAV}, {OUT_SHARES}, {OUT_SUMMARY}")


if __name__ == "__main__":
    main()
