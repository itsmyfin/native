"""
ESG vs Plain ETF Research Note — Chart Generator
=================================================
Generates 5 publication-ready charts for any ESG/plain ETF benchmark pair.

USAGE
-----
1. Fill in FIRM_COLORS with your 15 hex codes (slots are labelled).
2. Call generate_all_charts(etf_data, output_dir) with your data dict.
3. One PNG per chart is saved to output_dir, named by benchmark.

Run a single benchmark:
    python esg_charts.py

Run multiple benchmarks in a loop:
    for benchmark, data in all_benchmarks.items():
        generate_all_charts(data, output_dir=f"output/{benchmark}")
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D

# ── 1. FIRM COLOR PALETTE ──────────────────────────────────────────────────
# Replace each placeholder with your actual hex codes.
# Naming convention is intentional — map to your brand guide as needed.

FIRM_COLORS = {
    # Primary series colors
    "esg":          "#PLACEHOLDER_01",   # ESG ETF — bars, lines
    "plain":        "#PLACEHOLDER_02",   # Plain ETF — bars, lines
    "accent":       "#PLACEHOLDER_03",   # Highlight / annotation
    "drift":        "#PLACEHOLDER_04",   # Mandate drift line (overlap %)
    "flows_pos":    "#PLACEHOLDER_05",   # Positive flow bars
    "flows_neg":    "#PLACEHOLDER_06",   # Negative flow bars
    "ref_line":     "#PLACEHOLDER_07",   # Reference / dashed lines

    # Semantic fill areas
    "outperform":   "#PLACEHOLDER_08",   # Shaded area — ESG ahead
    "underperform": "#PLACEHOLDER_09",   # Shaded area — ESG behind

    # Chart furniture
    "background":   "#PLACEHOLDER_10",   # Figure background
    "panel":        "#PLACEHOLDER_11",   # Axes background
    "grid":         "#PLACEHOLDER_12",   # Gridlines
    "axis_text":    "#PLACEHOLDER_13",   # Tick labels, axis titles
    "title_text":   "#PLACEHOLDER_14",   # Chart title
    "annotation":   "#PLACEHOLDER_15",   # Annotation text / callout boxes
}

# ── 2. TYPOGRAPHY & GLOBAL STYLE ──────────────────────────────────────────

FONT_FAMILY = "Verdana"
FONT_SIZE_BASE = 12
FONT_SIZE_TITLE = 13
FONT_SIZE_LABEL = 11
FONT_SIZE_TICK = 10
FONT_SIZE_ANNOTATION = 9

def apply_house_style(fig, axes):
    """Apply firm style to a figure and list of axes."""
    fig.patch.set_facecolor(FIRM_COLORS["background"])
    for ax in (axes if isinstance(axes, (list, tuple)) else [axes]):
        ax.set_facecolor(FIRM_COLORS["panel"])
        ax.tick_params(colors=FIRM_COLORS["axis_text"], labelsize=FONT_SIZE_TICK)
        ax.xaxis.label.set_color(FIRM_COLORS["axis_text"])
        ax.yaxis.label.set_color(FIRM_COLORS["axis_text"])
        ax.title.set_color(FIRM_COLORS["title_text"])
        for spine in ax.spines.values():
            spine.set_edgecolor(FIRM_COLORS["grid"])
            spine.set_linewidth(0.5)
        ax.grid(True, color=FIRM_COLORS["grid"], linewidth=0.5, linestyle="--", alpha=0.7)
        ax.set_axisbelow(True)

def make_fig(figsize=(10, 5.5)):
    fig, ax = plt.subplots(figsize=figsize)
    plt.rcParams.update({
        "font.family": FONT_FAMILY,
        "font.size": FONT_SIZE_BASE,
        "text.color": FIRM_COLORS["title_text"],
    })
    return fig, ax

def save_chart(fig, output_dir, filename):
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, filename)
    fig.savefig(path, dpi=150, bbox_inches="tight",
                facecolor=FIRM_COLORS["background"])
    plt.close(fig)
    print(f"  Saved: {path}")

# ── 3. DATA SCHEMA ─────────────────────────────────────────────────────────
#
# Pass a dict with this structure to generate_all_charts():
#
# etf_data = {
#     "benchmark":   "S&P 500",          # string — used in titles and filenames
#
#     # Chart 1 — annualised returns
#     "returns": {
#         "labels":  ["1 year", "3 year", "5 year"],
#         "esg":     [11.2, 9.4, 12.1],  # % annualised
#         "plain":   [13.8, 8.7, 11.3],
#     },
#
#     # Chart 2 — rolling 12m differential
#     "rolling_diff": {
#         "dates":   [...],               # list of date strings or datetime
#         "values":  [...],               # ESG minus plain, %
#         "annotations": [               # optional — list of (date_index, label)
#             (6,  "Tech rally"),
#             (14, "Energy rally"),
#         ],
#     },
#
#     # Chart 3 — sector weight delta
#     "sector_delta": {
#         "sectors": ["Technology", "Healthcare", "Energy", ...],
#         "deltas":  [4.2, 1.1, -3.8, ...],   # ESG minus plain, pp
#     },
#
#     # Chart 4 — mandate drift
#     "drift": {
#         "dates":        [...],
#         "overlap_pct":  [...],          # % ESG holdings also in plain ETF
#         "esg_energy":   [...],          # ESG energy sector weight %
#         "plain_energy": [...],          # Plain energy weight % (reference)
#     },
#
#     # Chart 5 — flows vs performance
#     "flows": {
#         "dates":        [...],
#         "net_flows":    [...],          # $bn, can be negative
#         "rolling_diff": [...],          # same differential series as chart 2
#     },
# }

# ── 4. CHART FUNCTIONS ─────────────────────────────────────────────────────

def chart1_returns(data, output_dir):
    """Grouped bar chart — annualised returns by horizon."""
    d = data["returns"]
    benchmark = data["benchmark"]
    labels = d["labels"]
    esg = np.array(d["esg"])
    plain = np.array(d["plain"])

    x = np.arange(len(labels))
    width = 0.35

    fig, ax = make_fig(figsize=(8, 5))
    bars_esg   = ax.bar(x - width/2, esg,   width, label="ESG ETF",
                        color=FIRM_COLORS["esg"],   zorder=3)
    bars_plain = ax.bar(x + width/2, plain, width, label="Plain ETF",
                        color=FIRM_COLORS["plain"], zorder=3,
                        hatch="//", edgecolor=FIRM_COLORS["background"])

    for bar in list(bars_esg) + list(bars_plain):
        ax.text(bar.get_x() + bar.get_width()/2,
                bar.get_height() + 0.15,
                f"{bar.get_height():.1f}%",
                ha="center", va="bottom",
                fontsize=FONT_SIZE_ANNOTATION,
                color=FIRM_COLORS["axis_text"],
                fontfamily=FONT_FAMILY)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.set_title(f"Annualised returns — {benchmark} ESG vs plain",
                 fontsize=FONT_SIZE_TITLE, fontfamily=FONT_FAMILY, pad=10)
    ax.set_ylabel("Annualised return", fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY)

    legend = ax.legend(fontsize=FONT_SIZE_LABEL,
                       facecolor=FIRM_COLORS["panel"],
                       edgecolor=FIRM_COLORS["grid"],
                       labelcolor=FIRM_COLORS["axis_text"])
    apply_house_style(fig, ax)
    fig.tight_layout()
    save_chart(fig, output_dir, f"01_returns_{benchmark.replace(' ', '_')}.png")


def chart2_rolling_diff(data, output_dir):
    """Line chart — rolling 12m return differential with shaded regions."""
    d = data["rolling_diff"]
    benchmark = data["benchmark"]
    dates = d["dates"]
    values = np.array(d["values"])
    annotations = d.get("annotations", [])

    fig, ax = make_fig(figsize=(11, 5))
    ax.axhline(0, color=FIRM_COLORS["ref_line"], linewidth=0.8, linestyle="-")
    ax.fill_between(dates, values, 0,
                    where=(values >= 0),
                    color=FIRM_COLORS["outperform"], alpha=0.35, label="ESG outperforming")
    ax.fill_between(dates, values, 0,
                    where=(values < 0),
                    color=FIRM_COLORS["underperform"], alpha=0.35, label="ESG underperforming")
    ax.plot(dates, values, color=FIRM_COLORS["esg"], linewidth=1.8, zorder=4)

    for idx, label in annotations:
        ax.annotate(label,
                    xy=(dates[idx], values[idx]),
                    xytext=(0, 14), textcoords="offset points",
                    ha="center", fontsize=FONT_SIZE_ANNOTATION,
                    fontfamily=FONT_FAMILY, color=FIRM_COLORS["annotation"],
                    arrowprops=dict(arrowstyle="-", color=FIRM_COLORS["annotation"],
                                   lw=0.8))

    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{'+' if v >= 0 else ''}{v:.1f}%"))
    ax.set_title(f"Rolling 12m return differential — {benchmark} ESG minus plain",
                 fontsize=FONT_SIZE_TITLE, fontfamily=FONT_FAMILY, pad=10)
    ax.set_ylabel("ESG minus plain (%)", fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY)

    legend_elements = [
        mpatches.Patch(facecolor=FIRM_COLORS["outperform"], label="ESG outperforming"),
        mpatches.Patch(facecolor=FIRM_COLORS["underperform"], label="ESG underperforming"),
        Line2D([0], [0], color=FIRM_COLORS["esg"], lw=1.8, label="Differential"),
    ]
    ax.legend(handles=legend_elements, fontsize=FONT_SIZE_LABEL,
              facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
              labelcolor=FIRM_COLORS["axis_text"])

    plt.xticks(rotation=35, ha="right", fontsize=FONT_SIZE_TICK, fontfamily=FONT_FAMILY)
    apply_house_style(fig, ax)
    fig.tight_layout()
    save_chart(fig, output_dir, f"02_rolling_diff_{benchmark.replace(' ', '_')}.png")


def chart3_sector_delta(data, output_dir):
    """Horizontal bar chart — sector weight delta (ESG minus plain)."""
    d = data["sector_delta"]
    benchmark = data["benchmark"]
    sectors = d["sectors"]
    deltas = np.array(d["deltas"])

    # Sort by delta so the chart reads cleanly
    order = np.argsort(deltas)
    sectors = [sectors[i] for i in order]
    deltas = deltas[order]

    colors = [FIRM_COLORS["esg"] if v >= 0 else FIRM_COLORS["plain"] for v in deltas]

    fig, ax = make_fig(figsize=(9, max(5, len(sectors) * 0.55)))
    bars = ax.barh(sectors, deltas, color=colors, zorder=3, height=0.6)

    for bar, val in zip(bars, deltas):
        offset = 0.05 if val >= 0 else -0.05
        ha = "left" if val >= 0 else "right"
        ax.text(val + offset, bar.get_y() + bar.get_height()/2,
                f"{'+' if val >= 0 else ''}{val:.1f}pp",
                va="center", ha=ha,
                fontsize=FONT_SIZE_ANNOTATION, fontfamily=FONT_FAMILY,
                color=FIRM_COLORS["axis_text"])

    ax.axvline(0, color=FIRM_COLORS["ref_line"], linewidth=0.8)
    ax.xaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{'+' if v > 0 else ''}{v:.0f}pp"))
    ax.set_title(f"Sector weight delta — {benchmark} ESG minus plain",
                 fontsize=FONT_SIZE_TITLE, fontfamily=FONT_FAMILY, pad=10)
    ax.set_xlabel("Weight difference (percentage points)",
                  fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY)
    ax.tick_params(axis="y", labelsize=FONT_SIZE_LABEL)

    legend_elements = [
        mpatches.Patch(facecolor=FIRM_COLORS["esg"],   label="ESG overweight"),
        mpatches.Patch(facecolor=FIRM_COLORS["plain"], label="ESG underweight"),
    ]
    ax.legend(handles=legend_elements, fontsize=FONT_SIZE_LABEL,
              facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
              labelcolor=FIRM_COLORS["axis_text"])

    apply_house_style(fig, ax)
    fig.tight_layout()
    save_chart(fig, output_dir, f"03_sector_delta_{benchmark.replace(' ', '_')}.png")


def chart4_mandate_drift(data, output_dir):
    """Dual-axis line chart — holdings overlap % and energy weight convergence."""
    d = data["drift"]
    benchmark = data["benchmark"]
    dates = d["dates"]

    fig, ax1 = make_fig(figsize=(11, 5.5))
    ax2 = ax1.twinx()

    l1, = ax1.plot(dates, d["overlap_pct"],
                   color=FIRM_COLORS["drift"], linewidth=2,
                   marker="o", markersize=4, label="Holdings overlap %")
    l2, = ax2.plot(dates, d["esg_energy"],
                   color=FIRM_COLORS["accent"], linewidth=2,
                   marker="s", markersize=4, label="ESG energy weight")
    l3, = ax2.plot(dates, d["plain_energy"],
                   color=FIRM_COLORS["ref_line"], linewidth=1.2,
                   linestyle="--", label="Plain energy weight (ref)")

    ax1.set_ylabel("Holdings overlap (%)",
                   fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY,
                   color=FIRM_COLORS["drift"])
    ax2.set_ylabel("Energy sector weight (%)",
                   fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY,
                   color=FIRM_COLORS["accent"])
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))
    ax1.tick_params(axis="y", labelcolor=FIRM_COLORS["drift"])
    ax2.tick_params(axis="y", labelcolor=FIRM_COLORS["accent"])

    ax1.set_title(f"Mandate drift — {benchmark} ESG ETF holdings & sector convergence",
                  fontsize=FONT_SIZE_TITLE, fontfamily=FONT_FAMILY, pad=10)

    lines = [l1, l2, l3]
    labels = [l.get_label() for l in lines]
    ax1.legend(lines, labels, fontsize=FONT_SIZE_LABEL,
               facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
               labelcolor=FIRM_COLORS["axis_text"])

    plt.xticks(rotation=35, ha="right", fontsize=FONT_SIZE_TICK, fontfamily=FONT_FAMILY)
    apply_house_style(fig, [ax1, ax2])
    ax2.set_facecolor(FIRM_COLORS["panel"])
    for spine in ax2.spines.values():
        spine.set_edgecolor(FIRM_COLORS["grid"])
        spine.set_linewidth(0.5)
    fig.tight_layout()
    save_chart(fig, output_dir, f"04_mandate_drift_{benchmark.replace(' ', '_')}.png")


def chart5_flows(data, output_dir):
    """Bar + line combo — net flows vs rolling return differential."""
    d = data["flows"]
    benchmark = data["benchmark"]
    dates = d["dates"]
    flows = np.array(d["net_flows"])
    diff  = np.array(d["rolling_diff"])

    bar_colors = [FIRM_COLORS["flows_pos"] if f >= 0
                  else FIRM_COLORS["flows_neg"] for f in flows]

    fig, ax1 = make_fig(figsize=(13, 5.5))
    ax2 = ax1.twinx()

    ax1.bar(dates, flows, color=bar_colors, zorder=3, width=0.6,
            label="Net flows ($bn)")
    ax1.axhline(0, color=FIRM_COLORS["ref_line"], linewidth=0.6)

    l2, = ax2.plot(dates, diff, color=FIRM_COLORS["drift"],
                   linewidth=2, zorder=4, label="Return differential (12m rolling)")

    ax1.set_ylabel("Net flows ($bn)", fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY)
    ax2.set_ylabel("Return differential (%)",
                   fontsize=FONT_SIZE_LABEL, fontfamily=FONT_FAMILY,
                   color=FIRM_COLORS["drift"])
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:.0f}bn"))
    ax2.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{'+' if v >= 0 else ''}{v:.1f}%"))
    ax2.tick_params(axis="y", labelcolor=FIRM_COLORS["drift"])
    ax2.axhline(0, color=FIRM_COLORS["drift"], linewidth=0.4, linestyle=":")

    ax1.set_title(f"ESG ETF net flows vs return differential — {benchmark}",
                  fontsize=FONT_SIZE_TITLE, fontfamily=FONT_FAMILY, pad=10)

    legend_elements = [
        mpatches.Patch(facecolor=FIRM_COLORS["flows_pos"], label="Inflows ($bn)"),
        mpatches.Patch(facecolor=FIRM_COLORS["flows_neg"], label="Outflows ($bn)"),
        Line2D([0], [0], color=FIRM_COLORS["drift"], lw=2, label="Return diff (RHS)"),
    ]
    ax1.legend(handles=legend_elements, fontsize=FONT_SIZE_LABEL,
               facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
               labelcolor=FIRM_COLORS["axis_text"])

    plt.xticks(rotation=35, ha="right", fontsize=FONT_SIZE_TICK, fontfamily=FONT_FAMILY)
    apply_house_style(fig, [ax1, ax2])
    ax2.set_facecolor(FIRM_COLORS["panel"])
    for spine in ax2.spines.values():
        spine.set_edgecolor(FIRM_COLORS["grid"])
        spine.set_linewidth(0.5)
    fig.tight_layout()
    save_chart(fig, output_dir, f"05_flows_{benchmark.replace(' ', '_')}.png")


# ── 5. MAIN ENTRY POINT ───────────────────────────────────────────────────

def generate_all_charts(etf_data, output_dir="output"):
    """Generate all 5 charts for a single benchmark pair."""
    bm = etf_data["benchmark"]
    print(f"\nGenerating charts for: {bm}")
    chart1_returns(etf_data, output_dir)
    chart2_rolling_diff(etf_data, output_dir)
    chart3_sector_delta(etf_data, output_dir)
    chart4_mandate_drift(etf_data, output_dir)
    chart5_flows(etf_data, output_dir)
    print(f"Done — {bm}")


# ── 6. SAMPLE DATA (replace with your real data) ──────────────────────────

if __name__ == "__main__":

    import pandas as pd

    dates_monthly = pd.date_range("2020-01-01", "2024-12-01", freq="MS").tolist()
    dates_annual  = [2020, 2021, 2022, 2023, 2024, 2025]

    sp500_data = {
        "benchmark": "SP500",
        "returns": {
            "labels": ["1 year", "3 year", "5 year"],
            "esg":    [11.2, 9.4, 12.1],
            "plain":  [13.8, 8.7, 11.3],
        },
        "rolling_diff": {
            "dates":  dates_monthly,
            "values": np.interp(range(len(dates_monthly)),
                                [0, 12, 18, 30, 36, 48, 59],
                                [-0.3, 3.5, 2.1, -4.2, -1.9, 1.4, 1.8]).tolist(),
            "annotations": [
                (18, "Tech rally"),
                (30, "Energy rally"),
            ],
        },
        "sector_delta": {
            "sectors": ["Technology", "Healthcare", "Utilities",
                        "Consumer Disc", "Industrials", "Financials",
                        "Materials", "Energy"],
            "deltas":  [4.2, 1.1, 0.4, 0.3, -0.2, -0.9, -1.4, -3.8],
        },
        "drift": {
            "dates":        dates_annual,
            "overlap_pct":  [71, 74, 77, 79, 82, 84],
            "esg_energy":   [1.2, 1.5, 2.1, 2.3, 2.6, 2.8],
            "plain_energy": [4.5, 4.5, 4.5, 4.5, 4.5, 4.5],
        },
        "flows": {
            "dates":        dates_monthly,
            "net_flows":    np.interp(range(len(dates_monthly)),
                                      [0, 6, 18, 24, 36, 48, 59],
                                      [3, 7, 18, 8, -4, 5, 9]).tolist(),
            "rolling_diff": np.interp(range(len(dates_monthly)),
                                      [0, 12, 18, 30, 36, 48, 59],
                                      [-0.3, 3.5, 2.1, -4.2, -1.9, 1.4, 1.8]).tolist(),
        },
    }

    msci_em_data = {
        "benchmark": "MSCI_EM",
        "returns": {
            "labels": ["1 year", "3 year", "5 year"],
            "esg":    [6.4, 4.1, 7.8],
            "plain":  [7.9, 4.8, 8.2],
        },
        "rolling_diff": {
            "dates":  dates_monthly,
            "values": np.interp(range(len(dates_monthly)),
                                [0, 10, 20, 32, 45, 59],
                                [0.2, 2.1, -1.8, -3.1, 0.8, 1.2]).tolist(),
            "annotations": [
                (20, "EM risk-off"),
                (32, "Commodity spike"),
            ],
        },
        "sector_delta": {
            "sectors": ["Technology", "Consumer Disc", "Healthcare",
                        "Communication", "Industrials", "Financials",
                        "Materials", "Energy"],
            "deltas":  [5.1, 1.8, 0.9, 0.4, -0.3, -1.2, -2.1, -4.6],
        },
        "drift": {
            "dates":        dates_annual,
            "overlap_pct":  [68, 71, 74, 76, 79, 81],
            "esg_energy":   [2.1, 2.4, 3.0, 3.3, 3.6, 3.9],
            "plain_energy": [6.2, 6.2, 6.2, 6.2, 6.2, 6.2],
        },
        "flows": {
            "dates":        dates_monthly,
            "net_flows":    np.interp(range(len(dates_monthly)),
                                      [0, 8, 18, 28, 40, 59],
                                      [1.2, 4.5, 9.1, 3.2, -1.8, 3.4]).tolist(),
            "rolling_diff": np.interp(range(len(dates_monthly)),
                                      [0, 10, 20, 32, 45, 59],
                                      [0.2, 2.1, -1.8, -3.1, 0.8, 1.2]).tolist(),
        },
    }

    # ── Run for all benchmarks ──
    all_benchmarks = [sp500_data, msci_em_data]

    for etf_data in all_benchmarks:
        generate_all_charts(etf_data, output_dir="esg_charts_output")
