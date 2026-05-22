"""
ESG vs Plain ETF Research Note — PDF Report Generator
======================================================
Generates a fully formatted research note PDF including:
  - Cover page with benchmark name and date
  - Numbered sections with headings and body text
  - All 5 charts embedded inline at the correct section
  - Footer with page numbers and disclaimer

DEPENDENCIES
    pip install reportlab matplotlib pandas numpy

USAGE
    python esg_report_pdf.py
    → writes  esg_report_SP500.pdf  (and one per benchmark in the loop)

CUSTOMISATION
    1. Replace FIRM_COLORS hex codes with your 15 brand colors
    2. Replace FIRM_NAME / DISCLAIMER with your firm text
    3. Swap sample data for your real DataFrames
"""

import os
import io
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
from datetime import date

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image,
    HRFlowable, PageBreak, Table, TableStyle, KeepTogether
)
from reportlab.platypus.flowables import HRFlowable


# ── 1. FIRM SETTINGS ──────────────────────────────────────────────────────

FIRM_NAME      = "Your Firm Name"
REPORT_TITLE   = "ESG vs Plain ETF: Performance, Exposure & Mandate Integrity"
REPORT_DATE    = date.today().strftime("%d %B %Y")
DISCLAIMER     = (
    "This document is for informational purposes only and is directed at "
    "institutional and professional clients. Past performance is not a reliable "
    "indicator of future results. The data presented is illustrative — replace "
    "with your proprietary datasets before distribution."
)

# ── 2. FIRM COLOR PALETTE ─────────────────────────────────────────────────
# Replace each placeholder with your actual hex code.

FIRM_COLORS = {
    "esg":          "#PLACEHOLDER_01",
    "plain":        "#PLACEHOLDER_02",
    "accent":       "#PLACEHOLDER_03",
    "drift":        "#PLACEHOLDER_04",
    "flows_pos":    "#PLACEHOLDER_05",
    "flows_neg":    "#PLACEHOLDER_06",
    "ref_line":     "#PLACEHOLDER_07",
    "outperform":   "#PLACEHOLDER_08",
    "underperform": "#PLACEHOLDER_09",
    "background":   "#PLACEHOLDER_10",
    "panel":        "#PLACEHOLDER_11",
    "grid":         "#PLACEHOLDER_12",
    "axis_text":    "#PLACEHOLDER_13",
    "title_text":   "#PLACEHOLDER_14",
    "annotation":   "#PLACEHOLDER_15",
}

# ReportLab color objects derived from the palette
def rl_color(hex_code):
    """Convert hex string to ReportLab HexColor."""
    return colors.HexColor(hex_code)

# ── 3. TYPOGRAPHY ─────────────────────────────────────────────────────────

FONT            = "Helvetica"          # swap to registered TTF if Verdana is embedded
FONT_BOLD       = "Helvetica-Bold"
PAGE_W, PAGE_H  = A4
MARGIN          = 20 * mm

def build_styles():
    """Return a dict of named ParagraphStyles."""
    c_title  = rl_color(FIRM_COLORS["title_text"])
    c_body   = rl_color(FIRM_COLORS["axis_text"])
    c_accent = rl_color(FIRM_COLORS["esg"])

    return {
        "cover_firm": ParagraphStyle(
            "cover_firm", fontName=FONT_BOLD, fontSize=10,
            textColor=c_body, alignment=TA_LEFT, spaceAfter=4
        ),
        "cover_title": ParagraphStyle(
            "cover_title", fontName=FONT_BOLD, fontSize=20,
            textColor=c_title, alignment=TA_LEFT,
            spaceAfter=8, leading=26
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub", fontName=FONT, fontSize=11,
            textColor=c_body, alignment=TA_LEFT, spaceAfter=4
        ),
        "section_num": ParagraphStyle(
            "section_num", fontName=FONT_BOLD, fontSize=9,
            textColor=c_accent, alignment=TA_LEFT, spaceAfter=2
        ),
        "section_head": ParagraphStyle(
            "section_head", fontName=FONT_BOLD, fontSize=13,
            textColor=c_title, alignment=TA_LEFT,
            spaceBefore=14, spaceAfter=4
        ),
        "body": ParagraphStyle(
            "body", fontName=FONT, fontSize=10,
            textColor=c_body, alignment=TA_LEFT,
            spaceAfter=8, leading=15
        ),
        "chart_caption": ParagraphStyle(
            "chart_caption", fontName=FONT, fontSize=8,
            textColor=c_body, alignment=TA_LEFT,
            spaceAfter=12, leading=12
        ),
        "disclaimer": ParagraphStyle(
            "disclaimer", fontName=FONT, fontSize=7,
            textColor=c_body, alignment=TA_LEFT, leading=10
        ),
        "exec_summary": ParagraphStyle(
            "exec_summary", fontName=FONT, fontSize=10,
            textColor=c_body, alignment=TA_LEFT,
            spaceAfter=6, leading=15,
            leftIndent=10, rightIndent=10,
            borderPadding=(8, 10, 8, 10),
        ),
        "bullet": ParagraphStyle(
            "bullet", fontName=FONT, fontSize=10,
            textColor=c_body, alignment=TA_LEFT,
            spaceAfter=4, leading=14,
            leftIndent=16, bulletIndent=6
        ),
    }


# ── 4. PAGE TEMPLATE (header / footer) ───────────────────────────────────

def make_page_template(firm_name, report_title, disclaimer_text):
    """Returns an onPage callback for header/footer on every page."""
    def on_page(canvas, doc):
        canvas.saveState()
        w, h = A4

        # ── Header rule
        canvas.setStrokeColor(rl_color(FIRM_COLORS["esg"]))
        canvas.setLineWidth(1.5)
        canvas.line(MARGIN, h - 14*mm, w - MARGIN, h - 14*mm)

        # ── Firm name top-left, report title top-right
        canvas.setFont(FONT_BOLD, 7)
        canvas.setFillColor(rl_color(FIRM_COLORS["title_text"]))
        canvas.drawString(MARGIN, h - 10*mm, firm_name.upper())
        canvas.setFont(FONT, 7)
        canvas.setFillColor(rl_color(FIRM_COLORS["axis_text"]))
        canvas.drawRightString(w - MARGIN, h - 10*mm, report_title)

        # ── Footer rule
        canvas.setStrokeColor(rl_color(FIRM_COLORS["grid"]))
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, 14*mm, w - MARGIN, 14*mm)

        # ── Page number
        canvas.setFont(FONT, 7)
        canvas.setFillColor(rl_color(FIRM_COLORS["axis_text"]))
        canvas.drawRightString(w - MARGIN, 9*mm, f"Page {doc.page}")

        # ── Disclaimer footer
        canvas.setFont(FONT, 6)
        canvas.drawString(MARGIN, 9*mm, disclaimer_text[:120] + "…")

        canvas.restoreState()
    return on_page


# ── 5. CHART GENERATORS (return PNG bytes via BytesIO) ────────────────────

def _apply_style(fig, axes):
    fig.patch.set_facecolor(FIRM_COLORS["background"])
    for ax in (axes if isinstance(axes, (list, tuple)) else [axes]):
        ax.set_facecolor(FIRM_COLORS["panel"])
        ax.tick_params(colors=FIRM_COLORS["axis_text"], labelsize=9)
        ax.xaxis.label.set_color(FIRM_COLORS["axis_text"])
        ax.yaxis.label.set_color(FIRM_COLORS["axis_text"])
        ax.title.set_color(FIRM_COLORS["title_text"])
        for spine in ax.spines.values():
            spine.set_edgecolor(FIRM_COLORS["grid"])
            spine.set_linewidth(0.5)
        ax.grid(True, color=FIRM_COLORS["grid"], linewidth=0.4,
                linestyle="--", alpha=0.7)
        ax.set_axisbelow(True)

def _to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=FIRM_COLORS["background"])
    plt.close(fig)
    buf.seek(0)
    return buf

def _make_fig(w=9, h=4.5):
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
    return plt.subplots(figsize=(w, h))


def chart1_bytes(d, benchmark):
    labels = d["labels"]
    esg    = np.array(d["esg"])
    plain  = np.array(d["plain"])
    x      = np.arange(len(labels))
    width  = 0.32

    fig, ax = _make_fig(8, 4)
    ax.bar(x - width/2, esg,   width, color=FIRM_COLORS["esg"],
           label="ESG ETF", zorder=3, linewidth=0)
    ax.bar(x + width/2, plain, width, color=FIRM_COLORS["plain"],
           label="Plain ETF", zorder=3, linewidth=0,
           hatch="//", edgecolor=FIRM_COLORS["background"])
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))
    ax.set_title(f"Annualised returns — {benchmark} ESG vs plain", pad=8)
    ax.legend(fontsize=8, facecolor=FIRM_COLORS["panel"],
              edgecolor=FIRM_COLORS["grid"], labelcolor=FIRM_COLORS["axis_text"])
    _apply_style(fig, ax)
    fig.tight_layout()
    return _to_bytes(fig)


def chart2_bytes(d, benchmark):
    dates  = d["dates"]
    values = np.array(d["values"])

    fig, ax = _make_fig(9, 4)
    ax.axhline(0, color=FIRM_COLORS["ref_line"], linewidth=0.8)
    ax.fill_between(range(len(dates)), values, 0, where=(values >= 0),
                    color=FIRM_COLORS["outperform"], alpha=0.35)
    ax.fill_between(range(len(dates)), values, 0, where=(values < 0),
                    color=FIRM_COLORS["underperform"], alpha=0.35)
    ax.plot(range(len(dates)), values, color=FIRM_COLORS["esg"],
            linewidth=1.6, zorder=4)

    for idx, label in d.get("annotations", []):
        ax.annotate(label, xy=(idx, values[idx]),
                    xytext=(0, 12), textcoords="offset points",
                    ha="center", fontsize=8, color=FIRM_COLORS["annotation"],
                    arrowprops=dict(arrowstyle="-", color=FIRM_COLORS["annotation"],
                                   lw=0.7))

    step = max(1, len(dates) // 10)
    ax.set_xticks(range(0, len(dates), step))
    ax.set_xticklabels([str(dates[i])[:7] for i in range(0, len(dates), step)],
                       rotation=35, ha="right")
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{'+' if v>=0 else ''}{v:.1f}%"))
    ax.set_title(f"Rolling 12m return differential — {benchmark}", pad=8)
    ax.set_ylabel("ESG minus plain")

    legend_elements = [
        mpatches.Patch(facecolor=FIRM_COLORS["outperform"], label="ESG outperforming"),
        mpatches.Patch(facecolor=FIRM_COLORS["underperform"], label="ESG underperforming"),
    ]
    ax.legend(handles=legend_elements, fontsize=8,
              facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
              labelcolor=FIRM_COLORS["axis_text"])
    _apply_style(fig, ax)
    fig.tight_layout()
    return _to_bytes(fig)


def chart3_bytes(d, benchmark):
    sectors = d["sectors"]
    deltas  = np.array(d["deltas"])
    order   = np.argsort(deltas)
    sectors = [sectors[i] for i in order]
    deltas  = deltas[order]
    bar_colors = [FIRM_COLORS["esg"] if v >= 0 else FIRM_COLORS["plain"] for v in deltas]

    fig, ax = _make_fig(8.5, max(4, len(sectors) * 0.5))
    ax.barh(sectors, deltas, color=bar_colors, height=0.55, zorder=3)
    ax.axvline(0, color=FIRM_COLORS["ref_line"], linewidth=0.8)
    ax.xaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{'+' if v>0 else ''}{v:.0f}pp"))
    ax.set_title(f"Sector weight delta — {benchmark} ESG minus plain", pad=8)
    ax.set_xlabel("Weight difference (pp)")
    ax.tick_params(axis="y", labelsize=9)

    legend_elements = [
        mpatches.Patch(facecolor=FIRM_COLORS["esg"],   label="ESG overweight"),
        mpatches.Patch(facecolor=FIRM_COLORS["plain"], label="ESG underweight"),
    ]
    ax.legend(handles=legend_elements, fontsize=8,
              facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
              labelcolor=FIRM_COLORS["axis_text"])
    _apply_style(fig, ax)
    fig.tight_layout()
    return _to_bytes(fig)


def chart4_bytes(d, benchmark):
    dates = d["dates"]
    fig, ax1 = _make_fig(9, 4.5)
    ax2 = ax1.twinx()

    l1, = ax1.plot(range(len(dates)), d["overlap_pct"],
                   color=FIRM_COLORS["drift"], linewidth=2,
                   marker="o", markersize=4, label="Holdings overlap %")
    l2, = ax2.plot(range(len(dates)), d["esg_energy"],
                   color=FIRM_COLORS["accent"], linewidth=2,
                   marker="s", markersize=4, label="ESG energy weight")
    l3, = ax2.plot(range(len(dates)), d["plain_energy"],
                   color=FIRM_COLORS["ref_line"], linewidth=1.2,
                   linestyle="--", label="Plain energy weight (ref)")

    ax1.set_xticks(range(len(dates)))
    ax1.set_xticklabels([str(d_)[:4] for d_ in dates], rotation=35, ha="right")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))
    ax1.tick_params(axis="y", labelcolor=FIRM_COLORS["drift"])
    ax2.tick_params(axis="y", labelcolor=FIRM_COLORS["accent"])
    ax1.set_ylabel("Holdings overlap (%)", color=FIRM_COLORS["drift"])
    ax2.set_ylabel("Energy weight (%)", color=FIRM_COLORS["accent"])
    ax1.set_title(f"Mandate drift — {benchmark} ESG ETF", pad=8)

    lines, labels = [l1, l2, l3], [l.get_label() for l in [l1, l2, l3]]
    ax1.legend(lines, labels, fontsize=8, facecolor=FIRM_COLORS["panel"],
               edgecolor=FIRM_COLORS["grid"], labelcolor=FIRM_COLORS["axis_text"])

    _apply_style(fig, [ax1, ax2])
    ax2.set_facecolor(FIRM_COLORS["panel"])
    for spine in ax2.spines.values():
        spine.set_edgecolor(FIRM_COLORS["grid"]); spine.set_linewidth(0.5)
    fig.tight_layout()
    return _to_bytes(fig)


def chart5_bytes(d, benchmark):
    dates  = d["dates"]
    flows  = np.array(d["net_flows"])
    diff   = np.array(d["rolling_diff"])
    bar_colors = [FIRM_COLORS["flows_pos"] if f >= 0
                  else FIRM_COLORS["flows_neg"] for f in flows]

    fig, ax1 = _make_fig(10, 4.5)
    ax2 = ax1.twinx()

    ax1.bar(range(len(dates)), flows, color=bar_colors, zorder=3,
            width=0.65, label="Net flows ($bn)")
    ax1.axhline(0, color=FIRM_COLORS["ref_line"], linewidth=0.5)
    ax2.plot(range(len(dates)), diff, color=FIRM_COLORS["drift"],
             linewidth=1.8, zorder=4, label="Return differential (RHS)")

    step = max(1, len(dates) // 10)
    ax1.set_xticks(range(0, len(dates), step))
    ax1.set_xticklabels([str(dates[i])[:7] for i in range(0, len(dates), step)],
                        rotation=35, ha="right")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:.0f}bn"))
    ax2.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{'+' if v>=0 else ''}{v:.1f}%"))
    ax2.tick_params(axis="y", labelcolor=FIRM_COLORS["drift"])
    ax1.set_ylabel("Net flows ($bn)")
    ax2.set_ylabel("Return differential (%)", color=FIRM_COLORS["drift"])
    ax1.set_title(f"ESG ETF net flows vs return differential — {benchmark}", pad=8)

    legend_elements = [
        mpatches.Patch(facecolor=FIRM_COLORS["flows_pos"], label="Inflows"),
        mpatches.Patch(facecolor=FIRM_COLORS["flows_neg"], label="Outflows"),
        Line2D([0], [0], color=FIRM_COLORS["drift"], lw=1.8, label="Return diff (RHS)"),
    ]
    ax1.legend(handles=legend_elements, fontsize=8,
               facecolor=FIRM_COLORS["panel"], edgecolor=FIRM_COLORS["grid"],
               labelcolor=FIRM_COLORS["axis_text"])

    _apply_style(fig, [ax1, ax2])
    ax2.set_facecolor(FIRM_COLORS["panel"])
    for spine in ax2.spines.values():
        spine.set_edgecolor(FIRM_COLORS["grid"]); spine.set_linewidth(0.5)
    fig.tight_layout()
    return _to_bytes(fig)


# ── 6. PDF BUILDER ────────────────────────────────────────────────────────

USABLE_W = PAGE_W - 2 * MARGIN   # ~170mm on A4

def img_from_bytes(buf, width=USABLE_W):
    img = Image(buf, width=width, height=width * 0.48)
    img.hAlign = "LEFT"
    return img

def hr(thickness=0.5, color=None):
    return HRFlowable(
        width="100%",
        thickness=thickness,
        color=rl_color(color or FIRM_COLORS["grid"]),
        spaceAfter=4, spaceBefore=4
    )

def section(num, title, body_text, chart_buf=None,
            caption=None, styles=None, bullets=None):
    """Build a KeepTogether block for one note section."""
    s = styles
    items = [
        Paragraph(f"{'0' if num < 10 else ''}{num}", s["section_num"]),
        Paragraph(title, s["section_head"]),
        hr(),
        Paragraph(body_text, s["body"]),
    ]
    if bullets:
        for b in bullets:
            items.append(Paragraph(f"• {b}", s["bullet"]))
        items.append(Spacer(1, 6))
    if chart_buf:
        items.append(Spacer(1, 6))
        items.append(img_from_bytes(chart_buf))
        if caption:
            items.append(Paragraph(caption, s["chart_caption"]))
    items.append(Spacer(1, 10))
    return items


def build_pdf(etf_data, output_path):
    benchmark = etf_data["benchmark"]
    print(f"\nBuilding PDF: {output_path}")

    # ── Generate chart bytes
    c1 = chart1_bytes(etf_data["returns"],      benchmark)
    c2 = chart2_bytes(etf_data["rolling_diff"], benchmark)
    c3 = chart3_bytes(etf_data["sector_delta"], benchmark)
    c4 = chart4_bytes(etf_data["drift"],        benchmark)
    c5 = chart5_bytes(etf_data["flows"],        benchmark)

    # ── Document setup
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=22*mm, bottomMargin=20*mm,
        title=REPORT_TITLE,
        author=FIRM_NAME,
    )

    on_page = make_page_template(FIRM_NAME, REPORT_TITLE, DISCLAIMER)
    styles  = build_styles()
    story   = []

    # ── COVER PAGE ────────────────────────────────────────────────────────
    story += [
        Spacer(1, 30*mm),
        Paragraph(FIRM_NAME.upper(), styles["cover_firm"]),
        HRFlowable(width="100%", thickness=2,
                   color=rl_color(FIRM_COLORS["esg"]),
                   spaceAfter=10, spaceBefore=4),
        Paragraph(REPORT_TITLE, styles["cover_title"]),
        Spacer(1, 4*mm),
        Paragraph(f"Benchmark: {benchmark}", styles["cover_sub"]),
        Paragraph(f"Date: {REPORT_DATE}", styles["cover_sub"]),
        Spacer(1, 12*mm),
        HRFlowable(width="40%", thickness=0.5,
                   color=rl_color(FIRM_COLORS["grid"]),
                   spaceAfter=8, spaceBefore=0),
        Paragraph(
            "This note compares ESG-screened ETFs against their plain index "
            "counterparts across performance, risk-adjusted returns, sector "
            "exposure, and mandate integrity. Analysis covers 1y, 3y, and 5y "
            "horizons with rolling differential decomposition and a proprietary "
            "mandate drift score.",
            styles["body"]
        ),
        PageBreak(),
    ]

    # ── SECTION 01 — EXECUTIVE SUMMARY ───────────────────────────────────
    story += section(
        1, "Executive Summary",
        (
            "ESG-screened ETFs tracking the <b>{bm}</b> have delivered mixed "
            "results relative to plain index counterparts across the periods "
            "analysed. Outperformance on a 3y and 5y basis reflects structural "
            "sector tilts — particularly underweights in energy and materials — "
            "rather than genuine stock-selection alpha. On a 1y basis, the energy "
            "rally drove underperformance. Risk-adjusted metrics (Sharpe, "
            "drawdown) favour ESG in most regimes. Critically, holdings overlap "
            "with the plain index has risen materially, suggesting screen "
            "weakening over time."
        ).format(bm=benchmark),
        styles=styles,
        bullets=[
            "ESG ETFs outperformed on 3y and 5y; underperformed on 1y (energy-driven)",
            "Sector tilt — not ESG alpha — explains most of the return differential",
            "Holdings overlap has risen, indicating mandate drift",
            "Flows chased performance; crowding risk elevated in 2021-22",
            "Fee premium not fully justified by net-of-fee alpha on 1y horizon",
        ]
    )

    # ── SECTION 02 — METHODOLOGY ──────────────────────────────────────────
    story += section(
        2, "ETF Universe & Methodology",
        (
            "Analysis covers ESG-screened ETFs tracking the <b>{bm}</b> index "
            "against their plain equivalents. All ETFs use the same underlying "
            "index provider to isolate the effect of ESG screens rather than "
            "methodology differences. Returns are total return, net of fees, in "
            "local currency. Rebalance calendars are aligned. AUM minimum of "
            "$500m applied to exclude illiquid products. Holdings data sourced "
            "quarterly; sector weights monthly."
        ).format(bm=benchmark),
        styles=styles,
    )

    # ── SECTION 03 — RETURNS ─────────────────────────────────────────────
    story += section(
        3, "Annualised Return Comparison",
        (
            "The chart below shows annualised total returns for the ESG and plain "
            "ETF over 1, 3, and 5 year horizons. ESG outperformed on longer "
            "horizons driven by the structural underweight in energy, which "
            "lagged the broader market over 3-5 years. The 1-year picture "
            "reverses as the 2022 energy rally disproportionately benefited "
            "plain index holders."
        ),
        chart_buf=c1,
        caption=(
            "Chart 1: Annualised total returns (%), net of fees. "
            "Source: [your data provider]. Illustrative data — replace with actuals."
        ),
        styles=styles,
    )

    # ── SECTION 04 — RISK-ADJUSTED PERFORMANCE ────────────────────────────
    story += section(
        4, "Risk-Adjusted Performance",
        (
            "On a risk-adjusted basis, ESG ETFs display modestly superior Sharpe "
            "ratios over 3y and 5y horizons, driven by lower volatility from the "
            "energy underweight. Maximum drawdown during the 2022 selloff was "
            "shallower for ESG (-18.4% vs -21.2% for the plain index), providing "
            "partial downside protection. During the COVID drawdown, performance "
            "was broadly similar. Insert your Sharpe / vol / drawdown table here "
            "once actuals are loaded."
        ),
        styles=styles,
    )

    # ── SECTION 05 — ROLLING DIFFERENTIAL ────────────────────────────────
    story += section(
        5, "Rolling 12-Month Return Differential",
        (
            "The rolling 12-month return differential (ESG minus plain) reveals "
            "clear regime dependence. ESG outperformed strongly during the "
            "2020-21 tech-led rally when energy was depressed. The differential "
            "collapsed sharply into negative territory through 2022 as energy "
            "surged post-Ukraine. A gradual recovery began in 2023 and has "
            "continued into 2024-25. Annotated regime events are shown directly "
            "on the chart."
        ),
        chart_buf=c2,
        caption=(
            "Chart 2: Rolling 12-month return differential, ESG minus plain (%). "
            "Shaded green = ESG outperforming; shaded red = ESG underperforming. "
            "Illustrative data."
        ),
        styles=styles,
    )

    # ── SECTION 06 — SECTOR DELTA ─────────────────────────────────────────
    story += section(
        6, "Sector Weight Divergence",
        (
            "The sector weight delta chart is the single most important decomposition "
            "in this note. The majority of the long-run return differential between "
            "ESG and plain ETFs is explained by this chart rather than by individual "
            "security selection. Technology is the largest ESG overweight; energy "
            "is the largest underweight. Clients should treat ESG ETF performance "
            "relative to the plain index as primarily a sector rotation bet, not "
            "an ESG quality premium."
        ),
        chart_buf=c3,
        caption=(
            "Chart 3: Sector weight delta (ESG minus plain, percentage points). "
            "Blue = ESG overweight; orange = ESG underweight. Illustrative data."
        ),
        styles=styles,
    )

    # ── SECTION 07 — TRACKING ERROR (text only) ───────────────────────────
    story += section(
        7, "Tracking Error vs Parent Index",
        (
            "Tracking error of the ESG ETF relative to the plain parent index "
            "has averaged [X]bp annualised over the 5-year period, rising from "
            "[X]bp in 2020 to [X]bp in 2024. The direction of travel matters: "
            "rising tracking error implies growing active tilt from the ESG "
            "screen, while falling tracking error signals convergence — the ETF "
            "is increasingly replicating its plain counterpart despite the ESG "
            "label. Insert tracking error time series chart here from your data."
        ),
        styles=styles,
    )

    # ── SECTION 08 — FACTOR DECOMPOSITION (text only) ────────────────────
    story += section(
        8, "Factor Decomposition",
        (
            "A simple OLS regression of the ESG/plain return differential against "
            "standard factor returns (value, growth, quality, momentum, low-vol) "
            "attributes approximately [X]% of the explained variance to quality "
            "and [X]% to momentum. The residual — genuine ESG screen alpha — is "
            "statistically insignificant over the full 5-year window. This "
            "indicates that clients paying an ESG fee premium are largely "
            "purchasing accidental factor tilts available more cheaply elsewhere. "
            "Insert factor regression output table here."
        ),
        styles=styles,
    )

    # ── PAGE BREAK before drift section ──────────────────────────────────
    story.append(PageBreak())

    # ── SECTION 09-10-11 — MANDATE DRIFT ─────────────────────────────────
    story += section(
        9, "Mandate Drift: Holdings Overlap & Sector Convergence",
        (
            "This section presents the differentiated analytical angle of this "
            "note. Holdings overlap — the share of ESG ETF constituents also "
            "present in the plain index — has risen from <b>{ov_start}%</b> in "
            "{yr_start} to <b>{ov_end}%</b> in {yr_end}. Simultaneously, the "
            "ESG ETF's energy sector weight has crept toward the plain index "
            "level, from {e_start}% to {e_end}% (plain index: {plain}%). "
            "Both trends are consistent with screen softening, either through "
            "index methodology changes or portfolio drift between rebalances."
        ).format(
            ov_start=etf_data["drift"]["overlap_pct"][0],
            ov_end=etf_data["drift"]["overlap_pct"][-1],
            yr_start=etf_data["drift"]["dates"][0],
            yr_end=etf_data["drift"]["dates"][-1],
            e_start=etf_data["drift"]["esg_energy"][0],
            e_end=etf_data["drift"]["esg_energy"][-1],
            plain=etf_data["drift"]["plain_energy"][0],
        ),
        chart_buf=c4,
        caption=(
            "Chart 4: Holdings overlap % (left axis, purple) and energy sector "
            "weight % (right axis). Dashed line = plain index energy weight reference. "
            "Rising overlap and converging energy weight signal mandate drift."
        ),
        styles=styles,
    )

    # ── SECTION 12 — FLOWS ────────────────────────────────────────────────
    story += section(
        12, "Net Flows vs Return Differential",
        (
            "ESG ETF flows peaked approximately 1-2 quarters after peak "
            "outperformance — a classic performance-chasing pattern. This "
            "timing mismatch means a significant cohort of investors entered "
            "at the point of maximum valuation premium, subsequently experiencing "
            "the full magnitude of the 2022 underperformance. Current flow "
            "momentum is recovering alongside the improving differential. "
            "Crowding metrics from our proprietary dataset are available on "
            "request for individual ETF names."
        ),
        chart_buf=c5,
        caption=(
            "Chart 5: Monthly net flows ($bn, bars) vs rolling 12m return "
            "differential (%, line, RHS). Blue bars = inflows; red bars = outflows."
        ),
        styles=styles,
    )

    # ── SECTION 13 — FEES ─────────────────────────────────────────────────
    story += section(
        13, "Fee Comparison & Net-of-Fee Alpha",
        (
            "ESG ETFs in this category charge an average expense ratio of [X]bps "
            "versus [X]bps for plain equivalents — a premium of approximately "
            "[X]bps. On a 5-year basis, cumulative gross outperformance of "
            "[X]% is reduced to [X]% net of the fee premium. On a 1-year basis, "
            "the fee drag turns marginal gross outperformance negative. Clients "
            "should assess whether the mandate integrity — which this note shows "
            "is declining — justifies continued payment of the ESG premium. "
            "Insert fee table here."
        ),
        styles=styles,
    )

    # ── SECTION 14 — CONCLUSIONS ──────────────────────────────────────────
    story += section(
        14, "Takeaways & Client Implications",
        "Key conclusions for portfolio and risk teams:",
        styles=styles,
        bullets=[
            (
                "ESG outperformance is sector-driven, not screen-alpha. "
                "Clients should model it as a sector tilt, not a quality screen."
            ),
            (
                "Mandate integrity is declining. Holdings overlap has risen "
                f"significantly since {etf_data['drift']['dates'][0]}, reducing "
                "the differentiation clients are paying for."
            ),
            (
                "Fee premium is marginal on a 5y basis and negative on 1y. "
                "Review whether the ESG label justifies the cost."
            ),
            (
                "Flows are recovering but crowding risk remains elevated in "
                "tech-heavy ESG exposures. Monitor sector concentration."
            ),
            (
                "For clients requiring genuine ESG integrity, direct indexing "
                "or custom basket construction offers tighter screen control "
                "than off-the-shelf ETFs."
            ),
        ]
    )

    # ── DISCLAIMER PAGE ───────────────────────────────────────────────────
    story += [
        PageBreak(),
        Spacer(1, 20*mm),
        hr(thickness=1, color=FIRM_COLORS["grid"]),
        Spacer(1, 4),
        Paragraph("Important Disclosures", styles["section_head"]),
        Spacer(1, 4),
        Paragraph(DISCLAIMER, styles["disclaimer"]),
    ]

    # ── BUILD ─────────────────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f"  PDF written: {output_path}")


# ── 7. SAMPLE DATA & ENTRY POINT ─────────────────────────────────────────

if __name__ == "__main__":
    import pandas as pd

    dates_monthly = pd.date_range("2020-01-01", "2024-12-01", freq="MS").tolist()
    dates_annual  = [2020, 2021, 2022, 2023, 2024, 2025]

    sp500_data = {
        "benchmark": "S&P 500",
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
            "annotations": [(18, "Tech rally"), (30, "Energy rally")],
        },
        "sector_delta": {
            "sectors": ["Technology","Healthcare","Utilities",
                        "Consumer Disc","Industrials","Financials",
                        "Materials","Energy"],
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
        "benchmark": "MSCI EM",
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
            "annotations": [(20, "EM risk-off"), (32, "Commodity spike")],
        },
        "sector_delta": {
            "sectors": ["Technology","Consumer Disc","Healthcare",
                        "Communication","Industrials","Financials",
                        "Materials","Energy"],
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

    os.makedirs("esg_pdf_output", exist_ok=True)
    for etf_data in [sp500_data, msci_em_data]:
        bm = etf_data["benchmark"].replace(" ", "_")
        build_pdf(etf_data, f"esg_pdf_output/esg_note_{bm}.pdf")
