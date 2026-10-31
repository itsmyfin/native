#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Quarterly Thematic ETF Chartbook generator (HTML + PDF). Realized data only, no forecasts.
Usage:  python build_chartbook.py        (edit the CONFIG block below; nothing else needs changing)
Needs:  pandas, numpy, openpyxl; for PDF: playwright (+ `playwright install chromium`) or wkhtmltopdf."""
import os, re, sys, math, html as _h, subprocess, shutil, warnings
import numpy as np, pandas as pd
warnings.filterwarnings('ignore')
try:   # all openpyxl imports live here (module level) so no function ever imports them locally
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import BarChart, LineChart, ScatterChart, BubbleChart, Reference, Series
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import SeriesLabel, DataPoint
    from openpyxl.chart.data_source import StrRef
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.worksheet.properties import PageSetupProperties
    from openpyxl.utils import get_column_letter as CL
    HAVE_XL = True
except ImportError as _e:
    HAVE_XL = False; _XL_ERR = str(_e)

# ============================ CONFIG: EDIT HERE ============================
# ---------- 1. Input files ----------
MAIN_FILE = 'MainNew4.xlsx'            # ETF static data plus the monthly Flows_/AuM_ columns
MAIN_SHEET = 0                         # sheet name or position (0 = first sheet) in MAIN_FILE
METRICS_FILE = 'ETF_metrics.xlsx'      # returns, volatility, Sharpe and drawdown per ticker (decimals; drawdown negative)
METRICS_SHEET = 0                      # sheet name or position in METRICS_FILE
HOLD_FILE = 'Holdings-Thematics.xlsx'  # holdings workbook
HOLD_SHEET_THEMATIC = 'All Holdings'   # sheet holding the holdings of the thematic ETFs
HOLD_SHEET_SP500 = 'S&P500'            # sheet with S&P 500 reference holdings (concentration benchmark in the concentration exhibit)
HOLD_SHEET_WORLD = 'MSCI World'        # sheet with MSCI World reference holdings (concentration benchmark in the concentration exhibit)
ACWI_MONTHLY_FILE = 'ACWI_monthly.xlsx'  # optional file of monthly MSCI ACWI total returns for the correlation baseline; None = use the Broad-Market Equity Composite
ACWI_MONTHLY_SHEET = 0                 # sheet name or position in ACWI_MONTHLY_FILE
ACWI_DATE_COL = 'Date'                 # month-end date column in ACWI_MONTHLY_FILE
ACWI_RET_COL = 'Return'                # monthly return column in ACWI_MONTHLY_FILE
ACWI_RET_IS_PERCENT = False            # True if returns are stored as percent (1.2) instead of decimals (0.012)
# ---------- 2. Outputs ----------
OUT_DIR = 'output'                     # folder for all outputs (created if missing)
OUT_HTML = 'thematic_etf_chartbook.html'        # file name of the HTML report
OUT_PDF = 'thematic_etf_chartbook.pdf'          # file name of the PDF report
OUT_XLSX = 'thematic_etf_chartbook_data.xlsx'   # file name of the Excel workbook (one sheet per exhibit with data, text and native charts)
MAKE_HTML = True                       # write the HTML report
MAKE_PDF = False                       # write the PDF (made from the HTML; needs playwright or wkhtmltopdf)
MAKE_EXCEL = True                      # write the Excel workbook
SHOW_DATA = True                       # show the DATA bullet (the realized-data fact) under each exhibit and on each Excel sheet
SHOW_SCREEN = True                     # show the SCREEN bullet (the screening criterion or reference limit) under each exhibit and on each Excel sheet
SHOW_METHOD = True                     # show the METHOD bullet (how the statistic is calculated and how to read it)
# ---------- 3. Reporting period ----------
REPORT_DATE = '2026-09-30'             # quarter-end date of the report; QTD, YTD and 1Y windows are derived from it
# ---------- 4. Main file column names ----------
M_ISIN = 'ISIN'                        # unique ETF share class identifier (join key; holdings 'Instrument' maps to it)
M_TICKER = 'Ticker'                    # ETF ticker (joins to the metrics file)
M_NAME = 'Fund Name'                   # fund name (used for labels)
M_ASSET = 'Asset Class'                # asset class column (filtered to EQUITY)
M_TYPE = 'Fund Type'                   # wrapper type column (filtered to ETF_TYPE; ETC/ETN excluded)
M_STYLE = 'Active or Passive'          # Active / Passive flag
M_LAUNCH = 'Share Class Launch Date'   # launch date (drives the launch counts)
M_INDEX = 'Underlying Index Name'      # benchmark index tracked (groups passive ETFs by index)
M_TER = 'TER(%)'                       # total expense ratio in percent (0.15 = 15 bps)
M_DOM = 'Domicile'                     # United States or Europe
M_STRAT = 'Equity Strategy 1'          # strategy column used to split peer groups and find thematics
M_T1 = 'Theme 1'                       # broad level (Segment)
M_T2 = 'Theme 2'                       # specific level (Theme)
FLOW_PREFIX = 'Flows_'                 # prefix of monthly net flow columns (Flows_YYYY-MM-DD)
AUM_PREFIX = 'AuM_'                    # prefix of month-end AuM columns (AuM_YYYY-MM-DD)
THEME_COL = M_T2                       # level the analysis runs at: M_T2 (specific themes) or M_T1 (broad segments)
# ---------- 5. Metrics file column names (decimals; drawdown negative) ----------
X_TICKER = 'Ticker'                    # join key to the main file
X_QTD = 'QTD_Return'                   # quarter-to-date total return
X_QTD_S = 'QTD_Start_Date'             # QTD window start (must equal the prior quarter-end)
X_QTD_E = 'QTD_End_Date'               # QTD window end (must equal REPORT_DATE)
X_YTD = 'YTD_Return'                   # year-to-date total return
X_YTD_S = 'YTD_Start_Date'             # YTD window start (must equal the prior year-end)
X_YTD_E = 'YTD_End_Date'               # YTD window end (must equal REPORT_DATE)
X_1Y = '1Y_Return'                     # trailing 1-year total return
X_VOL = 'Vol_1Y'                       # trailing 1-year volatility
X_SHARPE = 'Sharpe_1Y'                 # trailing 1-year Sharpe ratio
X_DD = 'Max_Drawdown_1Y'               # trailing 1-year maximum drawdown (negative)
X_1Y_S = '1Y_Start_Date'               # 1Y window start (must equal 12 months before REPORT_DATE)
X_1Y_E = '1Y_End_Date'                 # 1Y window end (must equal REPORT_DATE)
# ---------- 6. Holdings column names (Instrument = ETF key, ISIN = stock) ----------
H_ETF = 'Instrument'                   # ETF identifier in the holdings file
H_STOCK = 'ISIN'                       # stock identifier in the holdings file
H_NAME = 'Holding Name'                # stock name
H_W = 'Percentage of Fund Assets'      # holding weight in percent
H_MCAP = 'Market_Cap'                  # stock market cap, USD mn
H_TURN = '20D_Turnover ($m)'           # stock 20-day average daily turnover, USD mn
H_DATE = 'Latest Filing Date'          # holdings as-of date (used to flag stale holdings)
HOLD_ETF_MAP_TO = M_ISIN               # main-file column that the holdings ETF key matches: M_ISIN or M_TICKER
# ---------- 7. Labels and filters ----------
EQUITY = 'Equity'                      # Asset Class value kept in the universe
ETF_TYPE = 'ETF'                       # Fund Type value kept in the universe
ACTIVE = 'Active'                      # value meaning an actively managed fund
PASSIVE = 'Passive'                    # value meaning an index-tracking fund
THEME_VALUE = 'Themes'                 # Equity Strategy 1 value that identifies thematic ETFs
PEER_GROUPS = ['Region/Country', 'Sectors', 'Themes', 'Factor/Style']   # Equity Strategy 1 values compared in the peer exhibit
BASELINE_PEER = 'Region/Country'       # peer group used as the flow hurdle and as the fallback correlation baseline
DOM_US = 'United States'               # Domicile value for US funds
DOM_EU = 'Europe'                      # Domicile value for European funds
ACWI_TICKER = 'ACWI US'                # ticker of the MSCI ACWI reference ETF in the metrics file
SPY_TICKER = 'SPY US'                  # ticker of the S&P 500 reference ETF in the metrics file
# ---------- 8. Analysis parameters ----------
RF = 0.0425                            # risk-free rate used in the Sharpe ratio (for the method text only; Sharpe values come from the metrics file)
TOP_N = 5                              # number of themes shown in the gaining, losing and flow polarization exhibits
MIN_THEME_AUM = 100.0                  # USD mn: themes below this current AuM are excluded from rankings
MIN_BASE_AUM = 100.0                   # USD mn: themes whose opening AuM (quarter or year start) is below this are excluded from the gaining/losing ranking
OUTLIER_NF = 1.0                       # normalized flows above this (1.0 = 100%) are marked with a dagger as a small-base warning
MIN_ETF_AUM = 50.0                     # USD mn: funds below this are treated as at risk of closure in the action bullets
THEME_CAP = 0.25                       # maximum weight of one theme within the thematic holdings (used in the gaining-themes action)
STREAK_QTRS = 8                        # maximum number of quarters looked back when counting consecutive quarters of inflows or outflows
PARTICIPATION = 0.25                   # share of daily turnover assumed sold per day when computing days-to-liquidate
DANGER_FLOAT = 10.0                    # liquidity stress zone: percent of market cap held by thematic ETFs above which a stock is flagged
DANGER_DTL = 5.0                       # liquidity stress zone: days-to-liquidate above which a stock is flagged
DTL_SANITY = 250                       # days-to-liquidate above this trigger a data-check warning (usually a units problem)
STALE_DAYS = 90                        # holdings older than this many days before REPORT_DATE are flagged as stale
CORR_MONTHS = 24                       # months of returns used in the correlation matrix
N_HEAT = 16                            # maximum number of themes in the performance heatmap
N_OVERLAP = 8                          # number of themes (largest ETF each) in the overlap heatmap
NEFF_MIN = 8.0                         # indices with effective constituents N.eff below this are treated as highly concentrated
OVERLAP_FLAG = 15.0                    # percent overlap above which two ETFs are treated as one position
CORR_CLUSTER = 0.75                    # correlation above which two themes are grouped into one risk bucket
DD_CAP = -0.35                         # drawdown deeper than this (-0.35 = -35%) triggers a position limit in the asymmetry action
DD_CAP_WEIGHT = 0.01                   # maximum portfolio weight per theme whose drawdown is deeper than DD_CAP
RET_DD_MIN = 1.5                       # minimum 1Y return divided by maximum drawdown for a theme to be preferred
GAP_FLAG = -0.05                       # flow-chasing gap (dollar-weighted minus time-weighted return) worse than this triggers staged entry
LOW_NEFF_WEIGHT = 0.02                 # position limit applied to funds tracking an index with N.eff below NEFF_MIN
ENTRY_TRANCHES = 3                     # minimum number of instalments recommended for themes flagged by GAP_FLAG
FEE_PREMIUM_MAX_BPS = 15              # maximum fee premium (TER, bps) over the passive median that is acceptable for an active fund
PROXY_TOL = 0.03                       # tolerance for the check of AuM/flow-derived returns against reported QTD returns
# ---------- 9. Report wording ----------
SOURCE = 'ETF Desk Research, Bloomberg, FactSet'   # source line shown under every exhibit and in the footer
LBL_GOOD = 'Resilient Outperformers'   # name of the favourable zone on the asymmetry map
LBL_STRESS = 'Liquidity Stress Zone'   # name of the stress zone in the ownership exhibit
LABEL_T1 = 'Segment'                   # report word for Theme 1 (the broad level)
LABEL_T2 = 'Theme'                     # report word for Theme 2 (the specific level)
# =========================================================================
NAVY, GREEN, RED, BLUE, GREY = '#0a2540', '#1a7f4b', '#b42318', '#7aa2c4', '#64748b'
FORBIDDEN = ['forecast', 'expect', 'will likely', 'price target', 'outlook', 'projected', 'predict', 'anticipate']
LOG = []
def log(m, lvl='INFO'): LOG.append((lvl, m)); print(f'[{lvl}] {m}')
def fail(m): print('[ERROR]', m); sys.exit(1)
def read(path, sheet, what):
    if not os.path.exists(path): fail(f'{what} file not found: {path}')
    try: return pd.read_excel(path, sheet_name=sheet)
    except ValueError as e: fail(f'{what}: cannot read sheet {sheet!r} ({e})')
def need(df, cols, what):
    miss = [c for c in cols if c not in df.columns]
    if miss: fail(f'{what}: missing columns {miss}. First columns found: {list(df.columns)[:25]}')
# ----- formatting (negatives in parentheses)
def _n(v, d, scale=1, suf=''):
    x = round(v * scale, d); t = f'{abs(x):,.{d}f}{suf}'; return '-' + t if x < 0 else t
def P(v, d=1): return 'n/a' if v is None or v != v else _n(v, d, 100, '%')
def PP(v, d=1): return 'n/a' if v is None or v != v else _n(v, d, 1, '%')
def PTS(v): return 'n/a' if v is None or v != v else _n(v, 1, 100, '%')       # difference of two returns, shown as %
def NUM(v, d=2): return 'n/a' if v is None or v != v else _n(v, d)
def M(v, d=0):                                                             # input in USD millions -> $mn / $bn / $tn
    if v is None or v != v: return 'n/a'
    a = abs(v); t = f'${a / 1e6:,.1f}tn' if a >= 1e6 else (f'${a / 1000:,.1f}bn' if a >= 1000 else f'${a:,.{d}f}mn'); return '-' + t if v < 0 and round(a, d) > 0 else t
def B(s): return f'<b>{s}</b>'
def esc(s): return _h.escape(str(s))
def short(s, n=30): s = str(s); return s if len(s) <= n else s[:n - 1] + '…'
def names(lst, k=4): lst = list(lst); return ', '.join(map(str, lst[:k])) + (f' +{len(lst) - k} more' if len(lst) > k else '') if lst else 'none'
def paren(l, k=4):
    l = list(l); return f' ({B(names(l, k))})' if l else ' (none today)'
# ============================ LOAD & VALIDATE ============================
m = read(MAIN_FILE, MAIN_SHEET, 'Main'); m.columns = [str(c).strip() for c in m.columns]
need(m, [M_ISIN, M_TICKER, M_NAME, M_ASSET, M_TYPE, M_STYLE, M_LAUNCH, M_INDEX, M_TER, M_DOM, M_STRAT, M_T1, M_T2], 'Main')
nd = m[M_ISIN].duplicated().sum()
if nd: log(f'{nd} duplicate ISIN rows dropped (first kept)', 'WARN'); m = m.drop_duplicates(M_ISIN)
def tcols(prefix):
    out = {}
    for c in m.columns:
        if c.startswith(prefix):
            try: out[c] = pd.Timestamp(c[len(prefix):])
            except Exception: log(f'column {c} ignored: date not parseable', 'WARN')
    return out
fc, ac = tcols(FLOW_PREFIX), tcols(AUM_PREFIX)
months = sorted(set(fc.values()) & set(ac.values()))
if not months: fail('No Flows_/AuM_ month columns found')
def wide(cmap):
    cols = [c for c in cmap if cmap[c] in months]
    w = m.set_index(M_ISIN)[cols].apply(pd.to_numeric, errors='coerce'); w.columns = [cmap[c] for c in cols]; return w[months]
FLW, AUMW = wide(fc), wide(ac)
REP = pd.Timestamp(REPORT_DATE)
if REP not in months:
    REP = max(d for d in months if d <= REP); log(f'Report date not in data; using {REP.date()}', 'WARN')
PQE = REP - pd.offsets.QuarterEnd(1); YS = pd.Timestamp(REP.year - 1, 12, 31); Y1 = REP - pd.offsets.MonthEnd(12)
for nm, d in (('prior quarter-end', PQE), ('prior year-end', YS), ('1Y start', Y1)):
    if d not in months: fail(f'{nm} {d.date()} is not a month-end in the Flows_/AuM_ columns')
Q_LBL = f'Q{(REP.month - 1) // 3 + 1} {REP.year}'; REPLBL = REP.strftime('%d %b %Y')
st = m.set_index(M_ISIN)[[M_TICKER, M_NAME, M_ASSET, M_TYPE, M_STYLE, M_LAUNCH, M_INDEX, M_TER, M_DOM, M_STRAT, M_T1, M_T2]].copy()
st = st[(st[M_ASSET] == EQUITY) & (st[M_TYPE] == ETF_TYPE)].copy()
log(f'Universe: {len(st)} equity ETF share classes (of {len(m)} rows; ETC/ETN and non-equity excluded)')
st['peer'] = st[M_STRAT].where(st[M_STRAT].isin(PEER_GROUPS), 'Other')
st['theme'] = st[THEME_COL].fillna('Unclassified').astype(str)
st['active'] = st[M_STYLE] == ACTIVE
st['launch'] = pd.to_datetime(st[M_LAUNCH], errors='coerce')
st['new_q'] = (st.launch > PQE) & (st.launch <= REP); st['new_ytd'] = (st.launch > YS) & (st.launch <= REP)
Fu, Au = FLW.loc[st.index].fillna(0), AUMW.loc[st.index]
def fsum(a, b): return Fu.loc[:, [c for c in months if a < c <= b]].sum(axis=1)
st['flow_q'], st['flow_ytd'], st['flow_1y'] = fsum(PQE, REP), fsum(YS, REP), fsum(Y1, REP)
st['aum_q0'], st['aum_y0'], st['aum_1y0'], st['aum'] = Au[PQE].fillna(0), Au[YS].fillna(0), Au[Y1].fillna(0), Au[REP].fillna(0)
# ----- metrics
x = read(METRICS_FILE, METRICS_SHEET, 'Metrics'); x.columns = [str(c).strip() for c in x.columns]
if 'QTD_Start_Date.1' in x.columns and X_YTD_S not in x.columns: x = x.rename(columns={'QTD_Start_Date.1': X_YTD_S}); log('Duplicate QTD_Start_Date header read as YTD_Start_Date')
need(x, [X_TICKER, X_QTD, X_QTD_S, X_QTD_E, X_YTD, X_YTD_S, X_YTD_E, X_1Y, X_VOL, X_SHARPE, X_DD, X_1Y_S, X_1Y_E], 'Metrics')
x = x.drop_duplicates(X_TICKER).set_index(X_TICKER)
for c in (X_QTD, X_YTD, X_1Y, X_VOL, X_SHARPE, X_DD): x[c] = pd.to_numeric(x[c], errors='coerce')
for c in (X_QTD, X_YTD, X_1Y, X_VOL):
    if x[c].abs().quantile(.99) > 5: log(f'{c} looks like percent units; divided by 100', 'WARN'); x[c] = x[c] / 100
if x[X_DD].median() > 0: log('Max drawdown was positive; sign flipped', 'WARN'); x[X_DD] = -x[X_DD]
def weq(col, d): return pd.to_datetime(x[col], errors='coerce').dt.normalize() == d
x['win_ok'] = weq(X_QTD_S, PQE) & weq(X_QTD_E, REP) & weq(X_YTD_S, YS) & weq(X_YTD_E, REP) & weq(X_1Y_S, Y1) & weq(X_1Y_E, REP)
bad = list(x.index[~x.win_ok])
if bad: log(f'{len(bad)} metrics rows fail the date-window assertion and are excluded from return/risk stats: {names(bad, 8)}', 'WARN')
xm = x.rename(columns={X_QTD: 'ret_q', X_YTD: 'ret_ytd', X_1Y: 'ret_1y', X_VOL: 'vol', X_SHARPE: 'sharpe', X_DD: 'dd'})[['ret_q', 'ret_ytd', 'ret_1y', 'vol', 'sharpe', 'dd', 'win_ok']]
st = st.join(xm, on=M_TICKER); st['ret_ok'] = st.win_ok.eq(True)
if not (st.ret_ok & (st.peer == THEME_VALUE)).any(): fail('No thematic ETF passed the metrics date-window assertion. Check the QTD/YTD/1Y start and end date columns against REPORT_DATE.')
log(f'Metrics matched for {st.ret_ok.sum()} of {len(st)} ETFs ({(st.ret_ok & (st.peer == THEME_VALUE)).sum()} thematic)')
def refrow(t):
    if t not in xm.index or not bool(xm.at[t, 'win_ok']): log(f'Reference {t} unavailable', 'WARN'); return None
    return xm.loc[t]
ACWI, SPY = refrow(ACWI_TICKER), refrow(SPY_TICKER)
# ----- proxy cross-check
def ret_series(ids):
    A, F = AUMW.loc[ids], FLW.loc[ids].fillna(0); prev = A.shift(1, axis=1); mask = (prev > 0) & A.notna()
    num = (A.where(mask, 0) - F.where(mask, 0)).sum(); den = prev.where(mask, 0).sum()
    return num / den.replace(0, np.nan) - 1
thm_ids = st.index[st.peer == THEME_VALUE]
qm = [c for c in months if PQE < c <= REP]
Rq = ((AUMW.loc[thm_ids, REP] - FLW.loc[thm_ids, qm].fillna(0).sum(axis=1)) / AUMW.loc[thm_ids, PQE].replace(0, np.nan) - 1)
chk_ = (Rq - st.loc[thm_ids, 'ret_q']).abs().dropna()
if len(chk_): log(f'Proxy vs reported QTD return: median abs gap {chk_.median() * 100:.2f}pp; {(chk_ > PROXY_TOL).sum()} of {len(chk_)} ETFs above {PROXY_TOL * 100:.0f}pp tolerance', 'INFO')
# ============================ AGGREGATION ============================
def wavg(df, col, key):
    d = df[df.ret_ok & df[col].notna() & (df.aum > 0)]
    return (d[col] * d.aum).groupby(d[key]).sum() / d.aum.groupby(d[key]).sum()
def agg(df, key):
    g = df.groupby(key).agg(n=(M_TICKER, 'size'), aum=('aum', 'sum'), aum_q0=('aum_q0', 'sum'), aum_y0=('aum_y0', 'sum'), aum_1y0=('aum_1y0', 'sum'),
                            flow_q=('flow_q', 'sum'), flow_ytd=('flow_ytd', 'sum'), flow_1y=('flow_1y', 'sum'))
    for a, b, c in (('ogr_q', 'flow_q', 'aum_q0'), ('ogr_ytd', 'flow_ytd', 'aum_y0'), ('ogr_1y', 'flow_1y', 'aum_1y0')): g[a] = g[b] / g[c].replace(0, np.nan)
    for c in ('ret_q', 'ret_ytd', 'ret_1y', 'vol', 'sharpe', 'dd'): g[c] = wavg(df, c, key)
    return g
stT = st[st.peer == THEME_VALUE]
peer = agg(st[st.peer != 'Other'], 'peer'); th = agg(stT, 'theme')
th['launch_q'] = stT.groupby('theme').new_q.sum(); th['launch_ytd'] = stT.groupby('theme').new_ytd.sum()
tfl = Fu.loc[stT.index].groupby(stT.theme).sum()
th['m_first'], th['m_last'] = tfl[qm[0]], tfl[qm[-1]]; th['trend'] = th.m_last - th.m_first
th['neg3'] = (tfl[qm] < 0).all(axis=1); th['pos3'] = (tfl[qm] > 0).all(axis=1)
if ACWI is not None:
    th['ex_acwi_1y'], th['ex_acwi_q'] = th.ret_1y - ACWI.ret_1y, th.ret_q - ACWI.ret_q
if SPY is not None: th['ex_spy_1y'] = th.ret_1y - SPY.ret_1y
ok_base = (th.aum_q0 >= MIN_BASE_AUM) & (th.aum_y0 >= MIN_BASE_AUM)
elig = th[(th.aum >= MIN_THEME_AUM) & ok_base].copy(); small_ = list(th.index[(th.aum >= MIN_THEME_AUM) & ~ok_base])
if small_: log(f'{len(small_)} themes excluded from the gaining/losing ranking because base AuM is below {MIN_BASE_AUM:,.0f}mn: {names(small_, 6)}')
if len(elig) < 2: log('Fewer than 2 themes pass the size screens; ranking falls back to all themes above the minimum AuM', 'WARN'); elig = th[th.aum >= MIN_THEME_AUM].copy()
nT = max(1, min(TOP_N, len(elig) // 2))
elig['score'] = (elig.ogr_q.rank(pct=True) + elig.ogr_ytd.rank(pct=True)) / 2
gain = elig.sort_values('score', ascending=False).head(nT); lose = elig.sort_values('score').head(nT)
tot_aum, tot_q, tot_ytd = th.aum.sum(), th.flow_q.sum(), th.flow_ytd.sum()
# ============================ HOLDINGS ============================
def hload(sheet, mapping=True):
    h = read(HOLD_FILE, sheet, f'Holdings[{sheet}]'); h.columns = [str(c).strip() for c in h.columns]
    need(h, [H_ETF, H_STOCK, H_W], f'Holdings[{sheet}]'); h[H_W] = pd.to_numeric(h[H_W], errors='coerce'); return h
def normw(h, key): h = h[h[H_STOCK].notna() & (h[H_W] > 0)].copy(); h['w'] = h[H_W] / h.groupby(key)[H_W].transform('sum') * 100; return h
h0 = hload(HOLD_SHEET_THEMATIC); need(h0, [H_NAME, H_MCAP, H_TURN, H_DATE], 'Holdings[thematic]')
k2i = {i: i for i in st.index} if HOLD_ETF_MAP_TO == M_ISIN else dict(zip(st[M_TICKER], st.index))
h0['etf'] = h0[H_ETF].astype(str).str.strip().map(k2i)
log(f'Holdings join: {h0.etf.notna().mean() * 100:.1f}% of rows matched to main file via {HOLD_ETF_MAP_TO} ({h0.etf.nunique()} ETFs)', 'INFO' if h0.etf.notna().mean() > .9 else 'WARN')
h = normw(h0[h0.etf.notna() & h0.etf.isin(stT.index)], 'etf'); h[H_DATE] = pd.to_datetime(h[H_DATE], errors='coerce')
st['hhi_etf'] = (h.w ** 2).groupby(h.etf).sum(); st['n_hold'] = h.groupby('etf').size(); st['filing'] = h.groupby('etf')[H_DATE].max()
st['stale'] = (REP - st.filing).dt.days > STALE_DAYS; stT = st[st.peer == THEME_VALUE]
log(f'{int(stT.hhi_etf.notna().sum())} of {len(stT)} thematic ETFs have holdings; {int(stT.stale.sum())} filings older than {STALE_DAYS} days')
cand = stT[stT.hhi_etf.notna()]; pas = cand[(cand[M_STYLE] == PASSIVE) & cand[M_INDEX].notna()]
rws = []
if len(pas):
    for ix, gp in pas.groupby(M_INDEX): r_ = gp.aum.idxmax(); rws.append(dict(label=ix, etf=r_, theme=st.at[r_, 'theme'], aum=gp.aum.sum(), n_etf=len(gp), N=st.at[r_, 'n_hold'], hhi=st.at[r_, 'hhi_etf']))
else:
    log('No passive thematic ETFs with index names; concentration shown at ETF level', 'WARN')
    for i, r_ in cand.iterrows(): rws.append(dict(label=r_[M_NAME], etf=i, theme=r_.theme, aum=r_.aum, n_etf=1, N=r_.n_hold, hhi=r_.hhi_etf))
conc = pd.DataFrame(rws)
if len(conc): conc['neff'] = 1e4 / conc.hhi
refs = {}
for lab, sh in (('S&P 500', HOLD_SHEET_SP500), ('MSCI World', HOLD_SHEET_WORLD)):
    try:
        hr = normw(hload(sh), H_ETF); hh = (hr.w ** 2).groupby(hr[H_ETF]).sum(); refs[lab] = (float(hr.groupby(H_ETF).size().median()), float(hh.median()))
    except SystemExit: raise
    except Exception as e: log(f'Reference sheet {sh} skipped: {e}', 'WARN')
# theme-level aggregate portfolio HHI
hh_ = h.merge(st[['aum', 'theme']], left_on='etf', right_index=True); hh_['wa'] = hh_.w * hh_.aum
gw = hh_.groupby(['theme', H_STOCK]).wa.sum(); gw = gw / gw.groupby(level=0).transform('sum') * 100
th['hhi'] = (gw ** 2).groupby(level=0).sum(); th['neff'] = 1e4 / th.hhi
# ownership
hh_['pos'] = hh_.w / 100 * hh_.aum
g = hh_.groupby(H_STOCK).agg(pos=('pos', 'sum'), n_etfs=('etf', 'nunique'), name=(H_NAME, 'first'), mcap=(H_MCAP, 'median'), turn=(H_TURN, 'median'))
g = g[(g.mcap > 0) & (g.turn > 0) & (g.pos > 0)].copy()
g['float'] = g.pos / g.mcap * 100; g['dtl'] = g.pos / (PARTICIPATION * g.turn); g['danger'] = (g['float'] > DANGER_FLOAT) & (g.dtl > DANGER_DTL)
if (g.pos / (PARTICIPATION * g.turn) > DTL_SANITY).any(): log(f'{int((g.pos / (PARTICIPATION * g.turn) > DTL_SANITY).sum())} stocks show days-to-liquidate above {DTL_SANITY:.0f}; check turnover units and ETF AuM', 'WARN')
if len(g):
    tp = hh_.groupby([H_STOCK, 'theme']).pos.sum().reset_index().sort_values('pos').groupby(H_STOCK).tail(1).set_index(H_STOCK).theme; g['theme'] = tp
# monthly series, correlation, flow-chasing
def irr_m(v0, fl, vend):
    n = len(fl); f = lambda i: v0 * (1 + i) ** n + sum(c * (1 + i) ** (n - k - 1) for k, c in enumerate(fl)) - vend
    lo, hi = -0.5, 0.5; flo, fhi = f(lo), f(hi)
    if flo * fhi > 0: return np.nan
    for _ in range(100):
        mid = (lo + hi) / 2; fm = f(mid)
        if flo * fm <= 0: hi = mid
        else: lo, flo = mid, fm
    return (lo + hi) / 2
m12 = [c for c in months if Y1 < c <= REP]; fcr = []
for t in th.index[th.aum_1y0 >= MIN_THEME_AUM]:
    ids = stT.index[stT.theme == t]; rs = ret_series(ids)[m12].dropna()
    if len(rs) < 6: continue
    tw = float(np.prod(1 + rs) - 1); i = irr_m(AUMW.loc[ids, Y1].sum(), FLW.loc[ids, m12].fillna(0).sum().tolist(), AUMW.loc[ids, REP].fillna(0).sum())
    if i == i: fcr.append(dict(theme=t, tw=tw, dw=(1 + i) ** len(m12) - 1))
fcr = pd.DataFrame(fcr)
if len(fcr): fcr['gap'] = fcr.dw - fcr.tw
# ============================ SVG HELPERS ============================
W = 900
def S(inner, h): return f'<svg viewBox="0 0 {W} {h:.0f}" xmlns="http://www.w3.org/2000/svg">{inner}</svg>'
def tx(x, y, s, a='start', sz=11, fill='#1e293b', wt='400', ex=''): return f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="{a}" font-size="{sz}" fill="{fill}" font-weight="{wt}" {ex}>{esc(s)}</text>'
def ln(x1, y1, x2, y2, c='#e2e8f0', w=1, d=''): return f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="{c}" stroke-width="{w}" {"stroke-dasharray=%s" % chr(34) + d + chr(34) if d else ""}/>'
def rc(x, y, w, h, fill, op=1, st_=''): return f'<rect x="{x:.1f}" y="{y:.1f}" width="{max(w, 0):.1f}" height="{max(h, 0):.1f}" fill="{fill}" fill-opacity="{op}" {st_}/>'
def ticks(lo, hi, n=5):
    if hi <= lo: hi = lo + 1
    raw = (hi - lo) / n; mag = 10 ** math.floor(math.log10(raw)); step = next((s * mag for s in (1, 2, 2.5, 5, 10) if s * mag >= raw), 10 * mag)
    t0, t1 = math.floor(lo / step + 1e-9) * step, math.ceil(hi / step - 1e-9) * step
    return [round(t0 + i * step, 10) for i in range(int(round((t1 - t0) / step)) + 1)]
def isn(v): return v is None or v != v
def legend(items, x, y):
    s = ''
    for i, (nm, c) in enumerate(items): s += rc(x + i * 150, y - 10, 11, 11, c) + tx(x + i * 150 + 16, y, nm, sz=11)
    return s
def hbars(labels, series, fmt, colors=None, left=210, right=90, tick_fmt=None, bh=15):
    colors = colors or [NAVY, BLUE, GREEN, RED]; k = len(series)
    vals = [v for _, vs in series for v in vs if not isn(v)]; lo = min([0] + vals); hi = max([0] + vals)
    tk = ticks(lo, hi, 5); lo, hi = min(lo, tk[0]), max(hi, tk[-1]); pw = W - left - right; X = lambda v: left + (v - lo) / (hi - lo) * pw
    gh = k * bh + 10; top = 30 if k > 1 else 10; hgt = top + len(labels) * gh + 30; s = ''
    for t in tk: s += ln(X(t), top - 4, X(t), top + len(labels) * gh, '#e2e8f0' if abs(t) > 1e-12 else NAVY, 1 if abs(t) > 1e-12 else 1.5) + tx(X(t), top + len(labels) * gh + 14, (tick_fmt or fmt)(t), 'middle', 10, GREY)
    if k > 1: s += legend([(nm, colors[j]) for j, (nm, _) in enumerate(series)], left, 14)
    for i, l in enumerate(labels):
        y0 = top + i * gh + 5; s += tx(left - 8, y0 + k * bh / 2 + 3, short(l, 34), 'end')
        for j, (nm, vs) in enumerate(series):
            v = vs[i]; y = y0 + j * bh
            if isn(v): s += tx(X(0) + 4, y + bh - 4, 'n/a', sz=10, fill=GREY); continue
            a, b = sorted((X(0), X(v))); c = colors[j] if k > 1 else (GREEN if v >= 0 else RED)
            s += rc(a, y, b - a, bh - 2, c) + (tx(b + 4, y + bh - 4, fmt(v), sz=10, wt='700') if v >= 0 else tx(a - 4, y + bh - 4, fmt(v), 'end', 10, wt='700'))
    return S(s, hgt)
def heat(rows, cols, mat, fmt, vmin, vmax, diverge=False, left=200, ch=26, top=44, cmax=None, blank=False):
    nc = len(cols); cw = min(90, (W - left - 10) / nc); s = ''
    for j, c in enumerate(cols): s += tx(left + j * cw + cw / 2, top - 10, short(c, 12), 'middle', 10, wt='700')
    for i, r in enumerate(rows):
        s += tx(left - 8, top + i * ch + ch / 2 + 4, short(r, 32), 'end')
        for j in range(nc):
            v = mat[i][j]
            if isn(v):
                if blank: continue
                s += rc(left + j * cw, top + i * ch, cw - 2, ch - 2, '#f1f5f9') + tx(left + j * cw + cw / 2, top + i * ch + ch / 2 + 4, 'n/a', 'middle', 10, GREY); continue
            if diverge: a = min(abs(v) / max(cmax[j] if cmax else max(abs(vmin), abs(vmax)), 1e-9), 1); fill = f'rgba({"26,127,75" if v >= 0 else "180,35,24"},{0.10 + 0.75 * a:.2f})'; tc = '#fff' if a > .75 else '#1e293b'
            else: a = min(max((v - vmin) / (vmax - vmin + 1e-12), 0), 1); fill = f'rgba(10,37,64,{0.06 + 0.9 * a:.2f})'; tc = '#fff' if a > .5 else '#1e293b'
            s += rc(left + j * cw, top + i * ch, cw - 2, ch - 2, fill) + tx(left + j * cw + cw / 2, top + i * ch + ch / 2 + 4, fmt(v), 'middle', 10, tc, '700')
    return S(s, top + len(rows) * ch + 16)
def dumbbell(rows, fmt, log_=False, left=210, right=200, leg=('A', 'B'), colb=RED):
    vs = [v for r in rows for v in (r[1], r[2]) if not isn(v)]
    if log_: lo, hi = 10 ** math.floor(math.log10(min(vs))), 10 ** math.ceil(math.log10(max(vs))); tk = [10 ** k for k in range(int(math.log10(lo)), int(math.log10(hi)) + 1)]; T_ = math.log10
    else: tk = ticks(min(vs), max(vs), 5); lo, hi = tk[0], tk[-1]; T_ = lambda v: v
    pw = W - left - right; X = lambda v: left + (T_(min(max(v, lo), hi)) - T_(lo)) / (T_(hi) - T_(lo)) * pw; rh = 30; top = 30; s = ''
    for t in tk: s += ln(X(t), top - 6, X(t), top + len(rows) * rh, '#e2e8f0') + tx(X(t), top + len(rows) * rh + 14, fmt(t), 'middle', 10, GREY)
    s += f'<circle cx="{left}" cy="12" r="5" fill="#fff" stroke="{NAVY}" stroke-width="2"/>' + tx(left + 10, 16, leg[0]) + f'<circle cx="{left + 190}" cy="12" r="5" fill="{colb}"/>' + tx(left + 200, 16, leg[1])
    for i, r in enumerate(rows):
        y = top + i * rh + 14; s += tx(left - 8, y + 4, short(r[0], 34), 'end') + ln(X(r[1]), y, X(r[2]), y, '#94a3b8', 2)
        s += f'<circle cx="{X(r[1]):.1f}" cy="{y}" r="6" fill="#fff" stroke="{NAVY}" stroke-width="2"/><circle cx="{X(r[2]):.1f}" cy="{y}" r="6" fill="{r[4] if len(r) > 4 else colb}"/>' + tx(max(X(r[1]), X(r[2])) + 12, y + 4, r[3], sz=10)
    return S(s, top + len(rows) * rh + 28)
def bubble(pts, xt, yt, xf, yf, xl, yl, zones=(), vl=(), hl=(), xlog=False, ylog=False, h=400, left=70, rmax=28, lab_n=22):
    TX = (lambda v: math.log10(max(v, 1e-9))) if xlog else (lambda v: v); TY = (lambda v: math.log10(max(v, 1e-9))) if ylog else (lambda v: v)
    x0, x1, y0, y1 = TX(xt[0]), TX(xt[-1]), TY(yt[0]), TY(yt[-1]); pw = W - left - 30; top = 16; ph = h - top - 50
    X = lambda v: left + (min(max(TX(v), x0), x1) - x0) / (x1 - x0) * pw; Y = lambda v: top + (y1 - min(max(TY(v), y0), y1)) / (y1 - y0) * ph; s = ''
    for zx0, zx1, zy0, zy1, col, lab in zones: s += rc(X(zx0), Y(zy1), X(zx1) - X(zx0), Y(zy0) - Y(zy1), col, .09) + tx(X(zx1) - 6, Y(zy1) + 14, lab, 'end', 11, col, '700')
    for t in yt: s += ln(left, Y(t), left + pw, Y(t)) + tx(left - 6, Y(t) + 4, yf(t), 'end', 10, GREY)
    for t in xt: s += ln(X(t), top, X(t), top + ph) + tx(X(t), top + ph + 15, xf(t), 'middle', 10, GREY)
    for v, lab, c in vl: s += ln(X(v), top, X(v), top + ph, c, 1.5, '5,4') + tx(X(v) + 4, top + ph - 6, lab, sz=10, fill=c)
    for v, lab, c in hl: s += ln(left, Y(v), left + pw, Y(v), c, 1.5, '5,4') + tx(left + pw - 4, Y(v) - 4, lab, 'end', 10, c)
    mx = max([p[3] for p in pts if p[3]] + [1]); order = sorted(pts, key=lambda p: -(p[3] or 0))
    for k, (lab, xv, yv, sz, col) in enumerate(order):
        r = 6 if not sz else 5 + rmax * math.sqrt(sz / mx); s += f'<circle cx="{X(xv):.1f}" cy="{Y(yv):.1f}" r="{r:.1f}" fill="{col}" fill-opacity=".6" stroke="{col}"/>'
        if lab and (k < lab_n or col == RED): s += tx(X(xv) + r + 3, Y(yv) + 4, short(lab, 22), sz=10)
    s += tx(left + pw / 2, h - 8, xl, 'middle', 11, wt='700') + tx(14, top + ph / 2, yl, 'middle', 11, wt='700', ex=f'transform="rotate(-90 14 {top + ph / 2})"')
    return S(s, h)
def lines(xl, series, fmt, colors, left=80, right=170, h=300):
    vals = [v for _, vs in series for v in vs if not isn(v)]; tk = ticks(min(vals + [0]), max(vals + [0]), 5); lo, hi = tk[0], tk[-1]; top = 20; ph = h - top - 40
    X = lambda i: left + i * (W - left - right) / max(len(xl) - 1, 1); Y = lambda v: top + (hi - v) / (hi - lo) * ph; s = ''
    for t in tk: s += ln(left, Y(t), W - right, Y(t), '#e2e8f0' if t else NAVY, 1 if t else 1.5) + tx(left - 6, Y(t) + 4, fmt(t), 'end', 10, GREY)
    for i, l in enumerate(xl): s += tx(X(i), top + ph + 16, l, 'middle', 10, GREY)
    ends = []
    for j, (nm, vs) in enumerate(series):
        pts = ' '.join(f'{X(i):.1f},{Y(v):.1f}' for i, v in enumerate(vs) if not isn(v)); s += f'<polyline points="{pts}" fill="none" stroke="{colors[j % len(colors)]}" stroke-width="2.2"/>'; ends.append((Y(vs[-1]), nm, colors[j % len(colors)], vs[-1]))
    last = -99
    for y, nm, c, v in sorted(ends): y = max(y, last + 13); last = y; s += tx(W - right + 6, y + 4, f'{short(nm, 18)} {fmt(v)}', sz=10, fill=c, wt='700')
    return S(s, h)
def table(head, rows, cls='lg'): return f'<table class="{cls}"><tr>' + ''.join(f'<th>{esc(c)}</th>' for c in head) + '</tr>' + ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in r) + '</tr>' for r in rows) + '</table>'
def col(v, t): return f'<span style="color:{RED if (v == v and v < 0) else "inherit"}">{t}</span>'
# ============================ EXHIBITS ============================
EX = []
def chk(t):
    import re; z = re.sub('<[^>]+>', '', t).lower()
    for w in FORBIDDEN:
        if w in z: fail(f'Forward-looking word {w!r} found in auto text: {z[:120]}')
KEYMAP = [('Relative Normalized', 'peers'), ('Flow Polarization', 'pol'), ('Gaining Popularity', 'gain'), ('Losing Popularity', 'lose'), ('Flow Trend', 'trend'), ('Performance vs.', 'perf'), ('Asymmetry', 'asym'), ('Concentration Mirage', 'conc'), ('Liquidity Contagion', 'own'), ('Holdings Overlap', 'ovl'),
          ('Correlation Matrix', 'corr'), ('Product Lifecycle', 'launch'), ('Flow-Chasing', 'chase'), ('Appendix', 'lg')]
METHOD = {
 'peers': f'Normalized flow = net flows over the window ÷ AuM at the start of the window (prior quarter-end for QTD, prior year-end for YTD, 12 months earlier for 1Y), measured for each peer group in aggregate. Dividing by the base removes the size effect, so groups of different size compare directly. A group above its peers is gaining share of investor money whatever the market did; comparing the three windows shows whether momentum is building or fading. A listing share well above the AuM share points to many small funds.',
 'trend': 'Each line is the running total of monthly net flows (USD mn) since 1 January, summed across all ETFs in a theme. A rising line means steady inflows, a flat line little activity, and a falling line net redemptions. Compare recent slopes with earlier ones to see whether flows are accelerating or fading, and read the gap between the top and bottom lines as the dispersion between themes.',
 'pol': 'Bars show each theme\'s net flow for the quarter in USD mn (inflows to the right, outflows to the left), with its normalized flow beside the name. Absolute flows show where money moved in size; normalized flows show how large the move is against the theme\'s own base. Large bars with small percentages are big themes with routine flows, and small bars with large percentages are small themes with heavy relative activity. Gross rotation (inflows plus outflows) measures how much capital is moving between themes.',
 'gain': 'Themes above the minimum AuM are ranked on the average of their percentile ranks for QTD and YTD normalized flow; the top N are shown with flows, AuM-weighted returns, excess return over MSCI ACWI, N.eff and launches. Read across the row: strong flows with positive excess return and new launches point to broad support, while strong flows with weak returns suggest money is arriving ahead of performance. The arrow compares the last month of the quarter with the first.',
 'lose': 'The same ranking as the gainers table, taking the bottom N. Persistent redemptions across all months, negative excess return and few launches indicate a theme that investors are leaving; redemptions with strong returns point to profit-taking rather than rejection. Small funds in these themes carry the highest closure risk.',
 'perf': 'Returns are AuM-weighted averages of the fund-level QTD, YTD and 1Y total returns of each theme\'s ETFs. Excess return is the theme\'s 1Y return minus the 1Y return of MSCI ACWI (and S&P 500), a simple difference of returns. Green is positive and red negative, with shade showing size, and rows are sorted by 1Y excess return over MSCI ACWI. A theme beating MSCI ACWI but not S&P 500 was strong against the world but not against US large caps. Compare columns to see whether recent returns confirm or contradict the longer window.',
 'asym': f'The x axis is 1Y maximum drawdown (deepest peak-to-trough loss), the y axis is 1Y total return, and bubble size is AuM; all are AuM-weighted fund statistics. Dashed lines mark MSCI ACWI, so themes in the top-right quadrant beat the benchmark with a shallower drawdown and themes in the bottom-left did worse on both counts. Return ÷ |drawdown| is a quick efficiency check; themes with deep drawdowns deserve smaller position sizes.',
 'chase': 'Time-weighted return compounds the theme\'s monthly returns and ignores flows. Dollar-weighted return is the IRR of the theme\'s cash flows (opening AuM, monthly net flows, closing AuM), the return investors earned in aggregate. A negative gap means money arrived after gains or left before recoveries. Large negative gaps mark themes where timing risk is high and staged entry is more suitable.',
 'launch': 'Launches are ETFs whose share class launch date falls inside the quarter or year to date. Active share is the active fund count, AuM or net flow as a share of the total, by domicile. A high active share of launches with a low share of flows means supply is running ahead of demand. Comparing launches in gaining and losing themes shows whether issuers follow flows. New funds should be screened for AuM and track record before they are used.',
 'conc': 'Weights are rescaled to 100% across the stocks held. HHI is the sum of squared weights (between 10,000 ÷ N and 10,000) and N.eff = 10,000 ÷ HHI is the number of equally weighted stocks that would give the same concentration. The distance between N (stocks listed) and N.eff shows how few stocks drive returns: a low N.eff means the fund behaves like a handful of stocks however many it lists. Compare with broad indices to judge how concentrated a theme is.',
 'own': f'Each stock\'s position is the sum, across all thematic ETFs, of weight × ETF AuM. Share held = position ÷ the stock\'s market cap. Days to liquidate = position ÷ ({PARTICIPATION:.0%} participation × 20-day average daily turnover), the time to sell the whole thematic ETF position without exceeding that share of daily volume. Stocks above both thresholds fall in the {LBL_STRESS}, where simultaneous redemptions would weigh most. Market cap is total rather than free float, so share held is understated, and very large day counts usually signal a position that cannot be sold at normal volumes (or a data-unit issue worth checking).',
 'ovl': 'For each pair of ETFs, overlap = sum over stocks of the smaller of the two weights (0 to 100%). The heatmap compares the largest ETF in each theme. High overlap means two funds give the same exposure under different names, and low overlap means genuinely different holdings. Use it to avoid double counting and to treat highly overlapping funds as one position.',
 'corr': 'Pearson correlation of monthly returns over the trailing window. Theme returns are AuM-weighted and derived from AuM and flows. Values run from -1 to 1, and values near 1 mean the themes move together. Low correlations between themes, or between a theme and the broad market, show real diversification. With about two years of monthly data, treat each value as indicative rather than precise.',
 'lg': f'One row per theme with AuM, flows, normalized flows, returns, risk and concentration, all AuM-weighted fund statistics. Sharpe = (1Y return - {RF:.2%}) ÷ 1Y volatility. Use it as the lookup table behind the other exhibits and to check a theme on popularity, performance and risk at once.'}
TXT = {}
def add(title, body, fact, act):
    key = next(k for w, k in KEYMAP if w in title); meth = METHOD.get(key, ''); chk(fact + ' ' + act + ' ' + meth)
    TXT[key] = dict(title=re.sub('<[^>]+>', '', title), fact=re.sub('<[^>]+>', '', fact), act=re.sub('<[^>]+>', '', act), meth=meth)
    EX.append(dict(key=key, html=f'<section><h2>Exhibit @@N@@: {title}</h2>{body}<div class="s">Source: {SOURCE} | Data as of {REPLBL}</div>' + ('<ul>' + (f'<li><b class="t">DATA</b>{fact}</li>' if SHOW_DATA else '') + (f'<li><b class="t a">SCREEN</b>{act}</li>' if SHOW_SCREEN else '') + (f'<li><b class="t m">METHOD</b>{meth}</li>' if (meth and SHOW_METHOD) else '') + '</ul>' if (SHOW_DATA or SHOW_SCREEN or SHOW_METHOD) else '') + '</section>'))
def skip(title, why): add(title, f'<p class="note">Not available: {esc(why)}</p>', f'Not computed: {esc(why)}.', 'Supply the missing input and re-run the build.')
# --- 1 peers
pr = [p for p in PEER_GROUPS if p in peer.index]
if THEME_VALUE in peer.index:
    tp_ = peer.loc[THEME_VALUE]; oth = [p for p in pr if p != THEME_VALUE]; beat = [p for p in oth if tp_.ogr_ytd > peer.at[p, 'ogr_ytd']]
    sh_n = (st.peer == THEME_VALUE).mean(); sh_a = st[st.peer == THEME_VALUE].aum.sum() / max(st.aum.sum(), 1e-9)
    f1 = f'Themes posted normalized flows of {B(P(tp_.ogr_q))} QTD, {B(P(tp_.ogr_ytd))} YTD and {B(P(tp_.ogr_1y))} over 1Y, outgrowing {B(f"{len(beat)} of {len(oth)}")} peer groups on YTD normalized flows ({names(beat)}); thematics are {B(P(sh_n))} of equity ETF listings but {B(P(sh_a))} of AuM.'
    a1 = f'Hurdle: the {BASELINE_PEER} YTD normalized flow is {B(P(peer.at[BASELINE_PEER, "ogr_ytd"]) if BASELINE_PEER in peer.index else "n/a")}; a theme whose own YTD normalized flow exceeds it is attracting money faster than broad regional funds.'
    add('Relative Normalized Flows: Themes vs. Peer Groups (net flows as % of base AuM)', hbars(pr, [('QTD Norm. Flow', [peer.at[p, 'ogr_q'] for p in pr]), ('YTD Norm. Flow', [peer.at[p, 'ogr_ytd'] for p in pr]), ('1Y Norm. Flow', [peer.at[p, 'ogr_1y'] for p in pr])], P, [BLUE, '#3f6c93', NAVY], tick_fmt=lambda v: P(v, 0)), f1, a1)
else: skip('Relative Normalized Flows', 'no thematic ETFs found')
# --- 2 polarization
ins_, outs_ = th[th.flow_q > 0].nlargest(nT, 'flow_q'), th[th.flow_q < 0].nsmallest(nT, 'flow_q')
pol = pd.concat([ins_, outs_]).sort_values('flow_q', ascending=False); inn, outn = ins_.flow_q.sum(), outs_.flow_q.sum(); w3 = list(th.index[th.neg3 & (th.aum >= MIN_THEME_AUM)])
add(f'QTD Absolute Flow Polarization: Top {nT} Inflow vs. Top {nT} Outflow Themes ($ mn; label = normalized flow)', hbars([f'{t} · {P(th.at[t, "ogr_q"])}' for t in pol.index], [('QTD absolute flow', list(pol.flow_q))], M, left=250, tick_fmt=lambda v: M(v)),
    f'Top-{nT} inflows summed to {B(M(inn))} (led by {B(pol.index[0])} at {B(M(pol.flow_q.iloc[0]))}, normalized flow {P(th.at[pol.index[0], "ogr_q"])}) against outflows of {B(M(outn))} across {len(outs_)} redeeming themes' + (f' (worst: {B(outs_.index[0])}, normalized flow {P(th.at[outs_.index[0], "ogr_q"])})' if len(outs_) else '') + f': gross rotation {B(M(inn - outn))}.',
    (f'Watch-list criterion: net redemptions in every month of the quarter ({B(names(w3))}). Flows describe what other investors did and are not a signal on their own; the performance and risk exhibits provide the context.' if w3 else 'Watch-list criterion: net redemptions in every month of the quarter; no theme meets it this quarter. Flows describe what other investors did and are not a signal on their own.'))
# --- 3/4 scorecards
def NF(t):
    v = th.neff.get(t, np.nan); return f'{v:.1f}' if v == v else 'n/a'
def dag(v): return '†' if v == v and v > OUTLIER_NF else ''
def _card(df):
    return table([LABEL_T2, 'ETFs', 'AuM', 'QTD Abs. Flow', 'QTD Norm. Flow', 'YTD Abs. Flow', 'YTD Norm. Flow', f'{qm[0]:%b}→{qm[-1]:%b}', 'QTD Ret', '1Y Ret', '1Y vs MSCI ACWI', 'N.eff', 'QTD Launches'],
                 [[esc(t), int(r.n), M(r.aum), col(r.flow_q, M(r.flow_q)), col(r.ogr_q, P(r.ogr_q) + dag(r.ogr_q)), col(r.flow_ytd, M(r.flow_ytd)), col(r.ogr_ytd, P(r.ogr_ytd) + dag(r.ogr_ytd)),
                   f'<span style="color:{GREEN if r.trend > 0 else RED}">{"▲" if r.trend > 0 else "▼"}</span>', col(r.ret_q, P(r.ret_q)), col(r.ret_1y, P(r.ret_1y)),
                   col(r.get('ex_acwi_1y', np.nan), PTS(r.get('ex_acwi_1y', np.nan))), NF(t), int(r.launch_q)] for t, r in df.iterrows()])
def card(df):
    flag = ((df.ogr_q > OUTLIER_NF) | (df.ogr_ytd > OUTLIER_NF)).any()
    return _card(df) + (f'<div class="s">† Normalized flow above {OUTLIER_NF:.0%} reflects a small base AuM; read it together with the absolute flow.</div>' if flag else '')
mA, mZ = qm[0].strftime('%B'), qm[-1].strftime('%B')
if len(gain):
    ld = gain.iloc[0]; lead_sh = ld.aum / max(tot_aum, 1e-9); nrise = int((gain.trend > 0).sum()); n_pos = int((gain.flow_q > 0).sum())
    pre = '' if n_pos == len(gain) else (f'Only {n_pos} of the {len(gain)} top-ranked themes attracted net inflows; ' if n_pos else 'No theme attracted net inflows this quarter; ')
    add(f'Themes Gaining Popularity: Top {nT} on Composite QTD/YTD Normalized Flow Rank', card(gain),
        f'{pre}{B(gain.index[0])} ranks first with QTD normalized flow {B(P(ld.ogr_q))}, YTD normalized flow {B(P(ld.ogr_ytd))}, 1Y return {B(P(ld.ret_1y))} and {B(int(ld.launch_q))} new QTD launches; {B(f"{nrise} of {len(gain)}")} of these themes had higher flow in {mZ} than in {mA}.',
        (f'Screening filter: gainers with net inflows in every month of the quarter show consistent demand{paren(gain.index[gain.pos3])}; the others show uneven flows. Reference limit: no single theme above {THEME_CAP:.0%} of thematic holdings.' if n_pos else 'Screening filter: none of the top-ranked themes shows net inflows, so flows give no case for adding to any theme; returns and risk provide the context.'))
    lt = lose.iloc[0]; sm_ = int(stT[stT.theme.isin(lose.index) & (stT.aum < MIN_ETF_AUM)].shape[0]); n_neg = int((lose.flow_q < 0).sum())
    pre = '' if n_neg == len(lose) else (f'Only {n_neg} of the {len(lose)} lowest-ranked themes saw net redemptions; ' if n_neg else 'No theme saw net redemptions this quarter; ')
    add(f'Themes Losing Popularity: Bottom {nT} on Composite QTD/YTD Normalized Flow Rank', card(lose),
        f'{pre}{B(lose.index[0])} ranks lowest with QTD normalized flow {B(P(lt.ogr_q))}, YTD normalized flow {B(P(lt.ogr_ytd))} and 1Y return {B(P(lt.ret_1y))}; {B(f"{int(lose.neg3.sum())} of {len(lose)}")} of these themes had redemptions in every month of the quarter.',
        (f'Watch-list criterion: redemptions in every month of the quarter{paren(lose.index[lose.neg3])}.' if lose.neg3.any() else 'Watch-list criterion: redemptions in every month of the quarter; no theme meets it this quarter.') + f' Funds below {M(MIN_ETF_AUM)} AuM run a risk of closure: {B(sm_)} ETFs in these themes are below that size.')
else:
    skip('Themes Gaining Popularity', 'too few themes pass the size screens'); skip('Themes Losing Popularity', 'too few themes pass the size screens')
# --- 5 flow trend
ms_ = [c for c in months if YS < c <= REP]; cum = tfl[ms_].cumsum(axis=1)
sel = list(gain.index[:3]) + list(lose.index[:3])
tot_m = tfl[ms_].sum(); bm, wm = tot_m.idxmax(), tot_m.idxmin(); turn = [t for t in elig.index if np.sign(elig.at[t, 'flow_q']) != np.sign(elig.at[t, 'flow_ytd']) and elig.at[t, 'flow_q'] != 0]
if sel:
    add('Flow Trend: Cumulative YTD Absolute Flows, Top Gaining vs. Losing Themes ($ mn)', lines([c.strftime('%b') for c in ms_], [(t, list(cum.loc[t])) for t in sel], M, [GREEN, '#4ea77a', '#8fd1b0', RED, '#d6685d', '#e8a29c']),
        f'Aggregate thematic flows peaked in {B(bm.strftime("%b"))} at {B(M(tot_m.max()))} and bottomed in {B(wm.strftime("%b"))} at {B(M(tot_m.min()))}; YTD absolute flows total {B(M(tot_ytd))} (QTD {B(M(tot_q))}).',
        (f"Review trigger: themes whose flows this quarter ran opposite to the year-to-date direction (inflows YTD but outflows QTD, or the reverse): {B(names(turn))}." if turn else "Review trigger: flows this quarter running opposite to the year-to-date direction; no theme meets it this quarter."))
if not sel: skip('Flow Trend', 'no gaining or losing themes to plot')
# --- 6 performance vs market (returns and excess returns in one heatmap)
use_a, use_s = ACWI is not None, SPY is not None
topT = th.sort_values('aum', ascending=False).head(N_HEAT)
topT = topT.sort_values('ex_acwi_1y' if use_a else 'ret_1y', ascending=False)
if len(topT):
    cols_ = ['QTD', 'YTD', '1Y'] + (['1Y vs MSCI ACWI'] if use_a else []) + (['1Y vs S&P 500'] if use_s else [])
    hr_, hm_ = [], []
    for t, r in topT.iterrows(): hr_.append(t); hm_.append([r.ret_q, r.ret_ytd, r.ret_1y] + ([r.ex_acwi_1y] if use_a else []) + ([r.ex_spy_1y] if use_s else []))
    pad = [np.nan] * (use_a + use_s)
    for nm_, rr in (('MSCI ACWI', ACWI), ('S&P 500', SPY)):
        if rr is not None: hr_.append(nm_); hm_.append([rr.ret_q, rr.ret_ytd, rr.ret_1y] + pad)
    arr = np.array(hm_, dtype=float); cmax = [float(np.nanmax(np.abs(arr[:, :3])))] * 3 + [float(np.nanmax(np.abs(arr[:, 3:]))) if arr.shape[1] > 3 else 1.0] * (arr.shape[1] - 3)
    b1, w1 = topT.ret_1y.idxmax(), topT.ret_1y.idxmin()
    if use_a:
        n_out = int((th.ex_acwi_1y > 0).sum()); n_sp = int((th.ex_spy_1y > 0).sum()) if use_s else None; best = topT.index[0]
        sh_ok = [t for t in th.index if th.at[t, 'ex_acwi_q'] > 0 and th.at[t, 'ex_acwi_1y'] > 0 and th.at[t, 'sharpe'] > ACWI.sharpe]
        f6 = f'Over 1Y, {B(b1)} returned {B(P(topT.at[b1, "ret_1y"]))} versus {B(w1)} at {B(P(topT.at[w1, "ret_1y"]))} (MSCI ACWI {B(P(ACWI.ret_1y))}, QTD {P(ACWI.ret_q)}); {B(f"{n_out} of {len(th)}")} themes beat MSCI ACWI' + (f' and {B(n_sp)} beat S&P 500' if n_sp is not None else '') + f', led by {B(best)} at {B(PTS(topT.at[best, "ex_acwi_1y"]))} excess return.'
        nsh = len(sh_ok); vb = 'qualifies' if nsh == 1 else 'qualify'
        a6 = (f"Screening filter: a theme qualifies when it beat MSCI ACWI over both the quarter and the year and its Sharpe ratio exceeds MSCI ACWI's ({NUM(ACWI.sharpe)}): {B(nsh)} {'theme' if nsh == 1 else 'themes'} {vb}{paren(sh_ok)}." if nsh else f"Screening filter: a theme qualifies when it beat MSCI ACWI over both the quarter and the year with a Sharpe ratio above MSCI ACWI's ({NUM(ACWI.sharpe)}); no theme qualifies this quarter.")
    else:
        f6 = f'Over 1Y, {B(b1)} returned {B(P(topT.at[b1, "ret_1y"]))} versus {B(w1)} at {B(P(topT.at[w1, "ret_1y"]))}; the MSCI ACWI reference is unavailable.'; a6 = 'Rank themes on 1Y return and review once the MSCI ACWI reference is available.'
    add(f'Performance vs. the Broader Market: Total Return and 1Y Excess Return over MSCI ACWI and S&P 500 (AuM-Weighted, Top {len(topT)} by AuM)', heat(hr_, cols_, hm_, lambda v: P(v), -1, 1, True, cmax=cmax, blank=True), f6, a6)
else: skip('Performance vs. the Broader Market', 'no theme has return data')
# --- 8 asymmetry
pts = []
for t, r in th.iterrows():
    if isn(r.ret_1y) or isn(r.dd): continue
    good = ACWI is not None and r.ret_1y > ACWI.ret_1y and r.dd > ACWI.dd; pts.append((t, r.dd, r.ret_1y, r.aum, GREEN if good else (RED if r.ret_1y < 0 else GREY)))
if pts:
    xs, ys = [p[1] for p in pts] + [0], [p[2] for p in pts] + [0]; xt, yt = ticks(min(xs) - .03, 0, 5), ticks(min(ys) - .03, max(ys) + .03, 6)
    z = [(ACWI.dd, 0, ACWI.ret_1y, yt[-1], GREEN, LBL_GOOD.upper() + ': beat MSCI ACWI with a shallower drawdown')] if ACWI is not None else []
    ratio = {p[0]: p[2] / abs(p[1]) for p in pts if p[1] < 0}; okr = [t for t, v in ratio.items() if v >= RET_DD_MIN]; deep = [p[0] for p in pts if p[1] < DD_CAP]
    bt = max(ratio, key=ratio.get); wt_ = min(pts, key=lambda p: p[2])
    add('The Asymmetry Map: 1Y Total Return vs. 1Y Max Drawdown (bubble = AuM)', bubble(pts, xt, yt, P, P, '1Y Max Drawdown', '1Y Total Return', z, [(ACWI.dd, 'MSCI ACWI', GREY)] if ACWI is not None else [], [(ACWI.ret_1y, 'MSCI ACWI', GREY)] if ACWI is not None else []),
        f'{B(bt)} delivered {B(P(th.at[bt, "ret_1y"]))} for a drawdown of {B(P(th.at[bt, "dd"]))} ({B(f"{ratio[bt]:.1f}x")} return per unit of drawdown), while {B(wt_[0])} returned {B(P(wt_[2]))} through a {B(P(wt_[1]))} drawdown.',
        f'Efficiency filter: 1Y return of at least {RET_DD_MIN:.1f}x maximum drawdown: ' + (f'{B(len(okr))} qualifying{paren(okr)}' if okr else 'no theme qualifies this quarter') + f'. Sizing reference: a position limit of {DD_CAP_WEIGHT:.0%} per theme applies where drawdown is deeper than {P(DD_CAP, 0)}{paren(deep)}.')
else: skip('The Asymmetry Map', 'no theme has return and drawdown data')
# --- 9 concentration
if len(conc):
    cs = conc.sort_values('hhi', ascending=False).head(12); rws = [(r.label, r.N, r.neff, f'N={int(r.N)} → N.eff {r.neff:.1f} (HHI {r.hhi:,.0f})') for r in cs.itertuples()]
    for lab, (nn, hhv) in refs.items(): rws.append((lab, nn, 1e4 / hhv, f'N={nn:,.0f} → N.eff {1e4 / hhv:.1f} (HHI {hhv:,.0f})', GREEN))
    lowN = conc[conc.neff < NEFF_MIN]; sh_l = lowN.aum.sum() / max(conc.aum.sum(), 1e-9); rf = '; '.join(f'{k} HHI {v[1]:,.0f}' for k, v in refs.items())
    add('The Concentration Mirage: Nominal vs. Effective Constituents of Thematic Indices (log scale)', dumbbell(rws, lambda v: f'{v:g}', True, leg=('Nominal constituents N', 'Effective constituents N.eff')),
        f'Median thematic benchmark HHI is {B(f"{conc.hhi.median():,.0f}")} (N.eff {B(f"{1e4 / conc.hhi.median():.1f}")}); {B(f"{len(lowN)} of {len(conc)}")} indices have N.eff below {NEFF_MIN:.0f}, tracked by {B(P(sh_l))} of index-tracking AuM; most concentrated: {B(cs.iloc[0].label)} (N.eff {B(f"{cs.iloc[0].neff:.1f}")}). References: {rf}.',
        f'Sizing reference: a position limit of {LOW_NEFF_WEIGHT:.0%} applies to funds tracking an index with N.eff below {NEFF_MIN:.0f} ({B(f"{len(lowN)} of {len(conc)}")} indices today); comparing their top holdings with core equity holdings shows hidden overlap.')
else: skip('The Concentration Mirage', 'no holdings matched to thematic ETFs')
# --- 10 ownership
if len(g):
    dz = g[g.danger].sort_values('float', ascending=False); top = g.sort_values('float', ascending=False).head(10)
    xl_ = [10 ** k for k in range(int(math.floor(math.log10(max(g.dtl.min(), 1e-2)))), int(math.ceil(math.log10(max(g.dtl.max(), 10)))) + 1)]
    yl_ = [10 ** k for k in range(int(math.floor(math.log10(max(g['float'].min(), 1e-3)))), int(math.ceil(math.log10(max(g['float'].max(), 20)))) + 1)]
    lab8 = set(dz.index[:8]); pts = [(r['name'] if i in lab8 else '', r.dtl, r['float'], None, RED if r.danger else GREY) for i, r in g.iterrows()]
    sc = bubble(pts, xl_, yl_, lambda v: f'{v:g}', lambda v: f'{v:g}%', 'Days-to-Liquidate (position ÷ 20% of 20D ADTV, log scale)', '% of Market Cap Held by Thematic ETFs (log scale)',
                [(DANGER_DTL, xl_[-1], DANGER_FLOAT, yl_[-1], RED, LBL_STRESS.upper())], [(DANGER_DTL, '5 days', RED)], [(DANGER_FLOAT, '10%', RED)], True, True, 380, lab_n=0)
    tb = table(['Stock', f'Main {LABEL_T2.lower()}', 'ETFs', 'ETF position', 'Market cap', '% held', 'Days to liquidate', 'Zone'], [[esc(short(r['name'], 24)), esc(r.get('theme', '')), int(r.n_etfs), M(r.pos), M(r.mcap), PP(r['float']), f'{r.dtl:,.1f}', f'<b style="color:{RED}">STRESS</b>' if r.danger else ''] for _, r in top.iterrows()])
    t10 = h.sort_values('w', ascending=False).groupby('etf').head(10); hit = t10[t10[H_STOCK].isin(set(dz.index))].etf.unique(); a_hit = st.loc[hit, 'aum'].sum() / max(stT.aum.sum(), 1e-9)
    gz = (f' {B(int(dz.theme.isin(gain.index).sum()))} of them are held mainly through gaining themes and {B(int(dz.theme.isin(lose.index).sum()))} through losing themes.' if len(gain) and len(dz) else '')
    add('Liquidity Contagion: Thematic ETF Ownership vs. Days-to-Liquidate (float proxy = total market cap)', sc + tb,
        f'{B(f"{len(dz)} of {len(g)}")} held stocks sit in the {LBL_STRESS} (above {DANGER_FLOAT:.0f}% held and over {DANGER_DTL:.0f} days); highest ownership is {B(top.iloc[0]["name"])} at {B(PP(top.iloc[0]["float"]))} of market cap and {B(f"{top.iloc[0].dtl:,.1f}")} days to liquidate.' + gz,
        f'Pre-trade check: a thematic ETF fails when any of its ten largest holdings is in the {LBL_STRESS}; {B(f"{len(hit)}")} thematic ETFs ({B(P(a_hit))} of thematic AuM) fail today.')
else: skip('Liquidity Contagion', 'no holdings with market cap and turnover')
# --- 11 overlap
reps = {}
for t in th.sort_values('aum', ascending=False).index:
    c_ = stT[(stT.theme == t) & stT.hhi_etf.notna()]
    if len(c_): reps[t] = c_.aum.idxmax()
    if len(reps) >= N_OVERLAP: break
if len(reps) >= 3:
    Wm = h.pivot_table(index=H_STOCK, columns='etf', values='w', aggfunc='sum').fillna(0); ks = list(reps); lab = [f'{st.at[reps[t], M_TICKER]} · {t}' for t in ks]
    O = [[100.0 if a == b else float(np.minimum(Wm[reps[a]], Wm[reps[b]]).sum()) for b in ks] for a in ks]
    pr_ = [(O[i][j], ks[i], ks[j]) for i in range(len(ks)) for j in range(i + 1, len(ks))]; hi_ = max(pr_); n15 = sum(1 for p in pr_ if p[0] > OVERLAP_FLAG); md = float(np.median([p[0] for p in pr_]))
    add(f'Holdings Overlap Heatmap: Largest ETF in Each of the Top {len(ks)} Themes (% shared weight)', heat(lab, [st.at[reps[t], M_TICKER] for t in ks], O, lambda v: f'{v:.0f}%', 0, max(30, hi_[0]), left=260),
        f'Median pairwise overlap is {B(PP(md))}, but {B(f"{n15} of {len(pr_)}")} pairs exceed {OVERLAP_FLAG:.0f}%; highest is {B(hi_[1])} / {B(hi_[2])} at {B(PP(hi_[0]))}.',
        f'Overlap rule: two funds with more than {OVERLAP_FLAG:.0f}% overlap count as one position under a combined limit, since holding both pays two fees for largely the same stocks ({B(n15)} of {len(pr_)} pairs meet it today).')
else: skip('Holdings Overlap Heatmap', 'fewer than 3 themes with holdings')
# --- 12 correlation
tops = list(th.sort_values('aum', ascending=False).index[:8]); R = pd.DataFrame({t: ret_series(stT.index[stT.theme == t]) for t in tops})
bname = 'Broad-Market Equity Composite'
if ACWI_MONTHLY_FILE and os.path.exists(ACWI_MONTHLY_FILE):
    a_ = read(ACWI_MONTHLY_FILE, ACWI_MONTHLY_SHEET, 'ACWI monthly'); need(a_, [ACWI_DATE_COL, ACWI_RET_COL], 'ACWI monthly')
    bs = pd.Series(pd.to_numeric(a_[ACWI_RET_COL]).values / (100 if ACWI_RET_IS_PERCENT else 1), index=pd.to_datetime(a_[ACWI_DATE_COL]) + pd.offsets.MonthEnd(0)).reindex(months); bname = 'MSCI ACWI'
else: bs = ret_series(st.index[st.peer == BASELINE_PEER])
R[bname] = bs; R = R.loc[months[-CORR_MONTHS:]]; C = R.corr(min_periods=8); cn = list(C.columns)
if len(tops) >= 3:
    prs = [(C.at[a, b], a, b) for i, a in enumerate(tops) for b in tops[i + 1:] if C.at[a, b] == C.at[a, b]]; avg = float(np.mean([p[0] for p in prs])); avb = float(C.loc[tops, bname].mean()); hi_, lo_ = max(prs), min(prs)
    clus = [f'{p[1]}/{p[2]}' for p in prs if p[0] > CORR_CLUSTER]; div = C.loc[tops, tops].apply(lambda r: (r.sum() - 1) / (len(tops) - 1)).idxmin()
    add(f'Realized Correlation Matrix: Top {len(tops)} Themes and {bname} (monthly, trailing {len(R)}M)', heat(cn, cn, C.values.tolist(), lambda v: f'{v:.2f}', -0.2, 1.0, left=240),
        f'Average pairwise theme correlation is {B(NUM(avg))} and average correlation to {bname} is {B(NUM(avb))}; highest pair {B(hi_[1] + " / " + hi_[2])} at {B(NUM(hi_[0]))}, lowest {B(lo_[1] + " / " + lo_[2])} at {B(NUM(lo_[0]))}.',
        f'Grouping rule: themes with correlation above {CORR_CLUSTER:.2f} form one risk bucket under a single limit{paren(clus)}; {B(div)} has the lowest average correlation to the other themes and adds the most diversification.')
else: skip('Realized Correlation Matrix', 'fewer than 3 themes')
# --- 13 launches
nq, nyt = int(stT.new_q.sum()), int(stT.new_ytd.sum()); aq = int((stT.new_q & stT.active).sum())
doms = [d for d in (DOM_US, DOM_EU) if d in set(stT[M_DOM])]
c1 = hbars(doms, [('Active', [int((stT.new_q & stT.active & (stT[M_DOM] == d)).sum()) for d in doms]), ('Passive', [int((stT.new_q & ~stT.active & (stT[M_DOM] == d)).sum()) for d in doms])], lambda v: f'{v:.0f}', [NAVY, BLUE], left=130, tick_fmt=lambda v: f'{v:.0f}')
def ashare(mask_fn, d):
    s = stT[stT[M_DOM] == d]; a = mask_fn(s); return a[0] / a[1] if a[1] else np.nan
labs = ['Active % of QTD launches', 'Active % of thematic AuM', 'Active % of QTD absolute flows', 'Active % of YTD absolute flows']
fns = [lambda s: ((s.new_q & s.active).sum(), s.new_q.sum()), lambda s: (s[s.active].aum.sum(), s.aum.sum()), lambda s: (s[s.active].flow_q.sum(), s.flow_q.sum() if s.flow_q.sum() > 0 else 0), lambda s: (s[s.active].flow_ytd.sum(), s.flow_ytd.sum() if s.flow_ytd.sum() > 0 else 0)]
c2 = hbars(labs, [(d, [ashare(f, d) for f in fns]) for d in doms], lambda v: P(v, 0), [NAVY, BLUE], left=250, tick_fmt=lambda v: P(v, 0))
nl = stT[stT.new_q].sort_values('aum', ascending=False).head(15)
tl = table(['Ticker', 'Fund', 'Theme', 'TER', 'Style', 'Domicile', 'AuM'], [[esc(r[M_TICKER]), esc(short(r[M_NAME], 36)), esc(r.theme), f'{r[M_TER]:.2f}%', esc(r[M_STYLE]), esc(r[M_DOM]), M(r.aum)] for _, r in nl.iterrows()]) if len(nl) else ''
lth = th.sort_values('launch_q', ascending=False); lt3 = [f'{t} ({int(v)})' for t, v in lth.launch_q.items() if v > 0][:3]
tera = stT[stT.active][M_TER].median(); terp = stT[~stT.active][M_TER].median(); smQ = int((stT.new_q & (stT.aum < MIN_ETF_AUM)).sum())
gl_txt = (f'{B(int(th.loc[gain.index, "launch_q"].sum()))} of {nq} QTD launches were in the top-{len(gain)} gaining themes vs. {B(int(th.loc[lose.index, "launch_q"].sum()))} in the bottom-{len(lose)} losing themes. ' if len(gain) and len(lose) and nq else '')
if nq == 0:
    f9 = f'No thematic ETFs launched in the quarter ({B(nyt)} year to date). Liquidations are not measured (delisted funds excluded from the file).'
    a9 = f'Screening filter: no new launches this quarter; the {M(MIN_ETF_AUM)} AuM and 12-month history criteria and the {FEE_PREMIUM_MAX_BPS:.0f} bps fee premium reference apply to the next batch.'
else:
    f9 = f'{B(nq)} thematic ETFs launched QTD ({B(nyt)} YTD), of which {B(f"{aq}")} active ({B(P(aq / nq))}); most launches: {B(names(lt3, 3))}. ' + gl_txt + 'Liquidations are not measured (delisted funds excluded from the file).'
    prem = (tera - terp) * 100
    fee = (f' Fee reference: an active fund is within tolerance when its TER is no more than {FEE_PREMIUM_MAX_BPS:.0f} bps above the passive equivalent (median today: active {B(f"{tera * 100:.0f} bps")} vs. passive {B(f"{terp * 100:.0f} bps")}, a premium of {B(f"{prem:.0f} bps")}).' if tera == tera and terp == terp else f' Fee reference: an active fund is within tolerance when its TER is no more than {FEE_PREMIUM_MAX_BPS:.0f} bps above the passive equivalent (no active/passive TER comparison available this quarter).')
    vb9 = 'is' if smQ == 1 else 'are'
    a9 = f'Screening filter: of {nq} QTD launches, {B(smQ)} {vb9} below {M(MIN_ETF_AUM)} AuM; funds under that size or with less than 12 months of history can be set aside.' + fee
add('Product Lifecycle: QTD Launches and the Active Migration, United States vs. Europe', c1 + c2 + tl, f9, a9)
# --- 14 flow chasing
if len(fcr):
    fs = fcr.sort_values('gap'); rws = [(r.theme, r.tw, r.dw, f'gap {PTS(r.gap)}') for r in fs.itertuples()]; wide_ = list(fs[fs.gap < GAP_FLAG].theme)
    add('The Flow-Chasing Penalty: Time-Weighted vs. Dollar-Weighted 1Y Return', dumbbell(rws, lambda v: P(v, 0), False, leg=('Time-weighted (fund)', 'Dollar-weighted (investors)')),
        f'Investors earned less than the fund in {B(f"{int((fcr.gap < 0).sum())} of {len(fcr)}")} themes; the average gap across all themes was {B(PTS(fcr.gap.mean()))}, with the largest shortfall in {B(fs.iloc[0].theme)} at {B(PTS(fs.iloc[0].gap))} and the most favourable in {B(fs.iloc[-1].theme)} at {B(PTS(fs.iloc[-1].gap))}.',
        (f'Staged-entry criterion: a gap worse than {PTS(GAP_FLAG)} between the investor and fund return points to timing risk{paren(wide_)}; entry in at least {ENTRY_TRANCHES} instalments is a common mitigant, and results are measured on dollar-weighted returns.' if wide_ else f'Staged-entry criterion: a gap worse than {PTS(GAP_FLAG)} between the investor and fund return; no theme meets it this quarter. Results are measured on dollar-weighted returns.'))
else: skip('The Flow-Chasing Penalty', 'insufficient monthly history')
# --- 15 league
lg = th.sort_values('aum', ascending=False).head(30)
rows_ = [[esc(t), int(r.n), M(r.aum), M(r.flow_q), P(r.ogr_q), P(r.ogr_ytd), P(r.ret_q), P(r.ret_ytd), P(r.ret_1y), P(r.vol), NUM(r.sharpe), P(r.dd), f'{r.hhi:,.0f}' if r.hhi == r.hhi else 'n/a', f'{r.neff:.1f}' if r.neff == r.neff else 'n/a'] for t, r in lg.iterrows()]
for nm_, rr in (('MSCI ACWI (ref.)', ACWI), ('S&P 500 (ref.)', SPY)):
    if rr is not None: rows_.append([f'<i>{nm_}</i>', '', '', '', '', '', P(rr.ret_q), P(rr.ret_ytd), P(rr.ret_1y), P(rr.vol), NUM(rr.sharpe), P(rr.dd), '', ''])
pos_s = list(th.index[(th.sharpe > (ACWI.sharpe if ACWI is not None else 0)) & (th.ogr_q > 0)])
add(f'Appendix: Theme League Table (Sharpe at Rf = {RF * 100:.2f}%; AuM-weighted fund statistics)', table([LABEL_T2, 'ETFs', 'AuM', 'QTD Abs. Flow', 'QTD Norm. Flow', 'YTD Norm. Flow', 'QTD', 'YTD', '1Y', 'Vol', 'Sharpe', 'Max DD', 'HHI', 'N.eff'], rows_),
    f'{B(f"{(th.sharpe < 0).sum()} of {len(th)}")} themes posted a negative 1Y Sharpe; only {B(f"{len(pos_s)}")} combine a Sharpe above {"MSCI ACWI" if ACWI is not None else "zero"} with positive QTD normalized flow.',
    f'Screening filter: a Sharpe ratio above the broad-market level together with positive QTD normalized flow; themes that meet both today{paren(pos_s)}.')
# ============================ HTML / PDF ============================
act_l = (stT[M_STYLE].eq(ACTIVE) & stT.new_q).sum() / max(stT.new_q.sum(), 1)
kp = [(M(tot_aum), 'Total Thematic AuM'), (M(tot_q), 'QTD Absolute Flows'), (M(tot_ytd), 'YTD Absolute Flows'), (f'{conc.hhi.median():,.0f}' if len(conc) else 'n/a', 'Median Thematic Index HHI'),
      (PP(g['float'].max()) if len(g) else 'n/a', 'Max Single-Stock Float Ownership'), (P(act_l, 0) if stT.new_q.sum() else 'n/a', 'Active Share of QTD Launches')]
CSS = """:root{color-scheme:light;box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;background:#eef1f5;font:14px/1.45 Georgia,serif;color:#1e293b;font-variant-numeric:tabular-nums}
.p{max-width:960px;margin:0 auto;background:#fff;padding:0 28px 24px}header{background:#0a2540;color:#fff;margin:0 -28px;padding:22px 28px}header h1{margin:0;font:700 24px Georgia,serif}header p{margin:4px 0 0;font:12px Arial,sans-serif;opacity:.8;letter-spacing:.04em;text-transform:uppercase}
.k{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:8px;margin:14px 0}.k div{border:1px solid #cbd5e1;border-top:3px solid #0a2540;padding:8px 10px}.k b{display:block;font:700 20px Arial,sans-serif;color:#0a2540}.k span{font:10px Arial,sans-serif;color:#64748b;text-transform:uppercase}
h3.sec{font:700 13px Arial,sans-serif;color:#fff;background:#0a2540;margin:26px 0 0;padding:6px 10px;break-after:avoid}.take{border:1px solid #cbd5e1;border-left:4px solid #1a7f4b;padding:8px 12px;margin:10px 0;font:12.5px/1.45 Georgia,serif}.take ul{margin:4px 0 0;padding-left:18px}section{margin:22px 0;break-inside:avoid;border-top:2px solid #0a2540;padding-top:8px}h2{font:700 15px Arial,sans-serif;color:#0a2540;margin:0 0 6px}svg{width:100%;height:auto;display:block;font-family:Arial,sans-serif}.s{font:10px Arial,sans-serif;color:#64748b;margin:2px 0 6px}
ul{margin:4px 0 0;padding-left:18px}li{margin:3px 0}li b.t{font:700 10px Arial,sans-serif;color:#fff;background:#0a2540;padding:1px 5px;margin-right:4px}li b.a{background:#1a7f4b}li b.m{background:#64748b}.note{font:11px Arial,sans-serif;background:#f1f5f9;padding:8px}
table{border-collapse:collapse;width:100%;font:10.5px Arial,sans-serif;margin-top:8px}th,td{border:1px solid #cbd5e1;padding:3px 5px;text-align:left}th{background:#e8edf3;color:#0a2540}.lg td:nth-child(n+2),.lg th:nth-child(n+2){text-align:right}
footer{font:10px/1.4 Arial,sans-serif;color:#64748b;border-top:2px solid #0a2540;margin-top:20px;padding-top:8px}@media print{body{background:#fff}.p{max-width:none;padding:0 8px}header{-webkit-print-color-adjust:exact;print-color-adjust:exact}svg,.k div{-webkit-print-color-adjust:exact;print-color-adjust:exact}}@page{size:A4;margin:10mm}"""
ORDER = [('Is money flowing into thematics?', ['peers']), ('Which themes are gaining or losing popularity?', ['pol', 'gain', 'lose', 'trend']),
         ('Did performance justify the flows?', ['perf', 'asym', 'chase']), ('Is the industry supplying more of it?', ['launch']),
         ('What risk sits under the hood?', ['conc', 'own', 'ovl']), ('What does it mean for the portfolio?', ['corr']), ('Appendix', ['lg'])]
byk = {e['key']: e for e in EX}; parts, n, sn = [], 0, 0; NUMS, SECS = {}, {}
for head, keys in ORDER:
    es = [byk[k] for k in keys if k in byk]
    if not es: continue
    sn += 1; parts.append(f'<h3 class="sec">{"Section " + str(sn) + " · " if head != "Appendix" else ""}{head}</h3>')
    for e in es: n += 1; NUMS[e['key']] = n; SECS[e['key']] = head; parts.append(e['html'].replace('@@N@@', str(n)))
tk = []
sg = agg(stT.assign(seg=stT[M_T1].fillna('Unclassified').astype(str)), 'seg'); sg = sg[sg.aum >= MIN_THEME_AUM]
if len(sg) >= 2:
    sg = sg.sort_values('ogr_q', ascending=False); a_, z_ = sg.iloc[0], sg.iloc[-1]
    tk.append(f'<b>By {LABEL_T1.lower()}:</b> {B(sg.index[0])} ranks highest on QTD normalized flow ({B(P(a_.ogr_q))}, {B(M(a_.flow_q))}; YTD {B(P(a_.ogr_ytd))}); {B(sg.index[-1])} ranks lowest ({B(P(z_.ogr_q))}, {B(M(z_.flow_q))}; YTD {B(P(z_.ogr_ytd))}), across {len(sg)} {LABEL_T1.lower()}s.')
if len(gain):
    g0 = gain.iloc[0]; npos = int((gain.flow_q > 0).sum())
    lead = (f'{B(gain.index[0])} leads on normalized flows (QTD {B(P(g0.ogr_q))}, YTD {B(P(g0.ogr_ytd))})' if npos else f'no theme attracted net inflows this quarter; {B(gain.index[0])} ranks first (QTD normalized flow {B(P(g0.ogr_q))})')
    tk.append(f'<b>Gaining:</b> {lead}; {B(f"{int(gain.pos3.sum())} of {len(gain)}")} top-ranked themes had inflows in every month of the quarter.')
if len(lose):
    l0 = lose.iloc[0]; nneg = int((lose.flow_q < 0).sum())
    wk = (f'{B(lose.index[0])} is weakest (QTD normalized flow {B(P(l0.ogr_q))}, 1Y return {B(P(l0.ret_1y))})' if nneg else f'no theme saw net redemptions this quarter; {B(lose.index[0])} ranks lowest (QTD normalized flow {B(P(l0.ogr_q))})')
    tk.append(f'<b>Losing:</b> {wk}' + (f'; {B(names(lose.index[lose.neg3], 3))} saw redemptions in every month of the quarter.' if lose.neg3.any() else '; no theme saw redemptions in every month of the quarter.'))
if ACWI is not None: tk.append(f'<b>Vs. the market:</b> {B(f"{int((th.ex_acwi_1y > 0).sum())} of {len(th)}")} themes beat MSCI ACWI over 1Y; {B(f"{nq}")} thematic ETFs launched QTD.')
if len(conc) and len(g): tk.append(f'<b>Risk:</b> median thematic index HHI is {B(f"{conc.hhi.median():,.0f}")}; {B(f"{int(g.danger.sum())}")} stocks sit in the {LBL_STRESS} (peak ownership {B(PP(g["float"].max()))} of market cap).')
for t_ in tk: chk(t_)
BODY = '<div class="take"><b>Key takeaways</b><ul>' + ''.join(f'<li>{t_}</li>' for t_ in tk) + '</ul></div>' + ''.join(parts)
BODY = re.sub(r"(?<=[A-Za-z])'(?=[a-z])", '’', BODY)   # typographic apostrophes in the visible text
dq = ''.join(f'<li>[{l}] {esc(t)}</li>' for l, t in LOG)
doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover"><title>Thematic ETF Monitor &amp; Strategy Chartbook {Q_LBL}</title><style>{CSS}</style></head><body><div class="p">
<header><p>ETF Desk Research · Equity Thematic Strategy · United States &amp; Europe</p><h1>Thematic ETF Monitor &amp; Strategy Chartbook</h1><p>{Q_LBL} (quarter ended {REPLBL}) · Realized data only · No forecasts</p></header>
<div class="k">{''.join(f'<div><b>{v}</b><span>{l}</span></div>' for v, l in kp)}</div>{BODY}
<section><h2>Data Quality Appendix (auto-generated)</h2><ul style="font:11px Arial,sans-serif">{dq}</ul></section>
<footer><b>Analyst Certification:</b> The research analyst(s) certify that the views expressed accurately reflect their personal views about the subject securities and that no part of their compensation was, is or will be directly or indirectly related to specific recommendations or views herein. <b>Disclosures:</b> For institutional / professional clients only (MiFID II professional client; FINRA institutional investor). Not retail marketing. Past performance is not indicative of future results. This is a retrospective analysis of realized data and contains no forecasts, price targets or investment advice. The portfolio rules shown are illustrative screening frameworks derived from realized data, not recommendations. Normalized flows = net flows ÷ base AuM; absolute flows = net flows in USD mn. HHI = Σwi²; N.eff = 10,000/HHI. Float ownership uses total market cap (free float not available) and days-to-liquidate assume 20% participation in 20-day average daily turnover. Theme statistics are AuM-weighted averages of fund-level statistics. Source: {SOURCE} | Data as of {REPLBL}</footer></div></body></html>"""
os.makedirs(OUT_DIR, exist_ok=True); hp = os.path.join(OUT_DIR, OUT_HTML)
if MAKE_HTML or MAKE_PDF: open(hp, 'w', encoding='utf-8').write(doc)   # a temporary HTML is written when only the PDF is requested
th.to_csv(os.path.join(OUT_DIR, 'theme_table.csv')); g.to_csv(os.path.join(OUT_DIR, 'stock_ownership.csv')); conc.to_csv(os.path.join(OUT_DIR, 'index_concentration.csv'), index=False); fcr.to_csv(os.path.join(OUT_DIR, 'flow_chasing.csv'), index=False)
print('HTML written:', hp) if MAKE_HTML else None
# ============================ EXCEL WORKBOOK (data + native charts, one sheet per exhibit) ============================
def write_excel(path):
    if not HAVE_XL: log(f'Excel export skipped: {_XL_ERR} (pip install openpyxl)', 'WARN'); return
    NAVYX = '0A2540'; HF = PatternFill('solid', fgColor=NAVYX); HFONT = Font(bold=True, color='FFFFFF'); WRAP = Alignment(wrap_text=True, vertical='top')
    PCT, MNF, INTF, DEC2, DEC1 = '0.0%', '#,##0.0', '#,##0', '0.00', '0.0'
    def py(v):
        if v is None: return None
        if isinstance(v, (float, np.floating)): return None if (v != v or v in (np.inf, -np.inf)) else float(v)
        if isinstance(v, np.integer): return int(v)
        if isinstance(v, np.bool_): return bool(v)
        if v is pd.NaT: return None
        if isinstance(v, pd.Timestamp): return v.to_pydatetime()
        return v
    wb = Workbook(); wb.remove(wb.active)
    SHORT = dict(peers='Peer flows', pol='Flow polarization', gain='Gaining themes', lose='Losing themes', trend='Flow trend', perf='Performance', asym='Asymmetry map',
                 chase='Flow chasing', launch='Launches', conc='Concentration', own='Stock ownership', ovl='Overlap', corr='Correlation', lg='League table')
    def mk_sheet(key):
        t = TXT.get(key, {}); ws = wb.create_sheet(f'Ex{NUMS[key]} {SHORT[key]}'[:31])
        ws['A1'] = f'Exhibit {NUMS[key]}: {t.get("title", SHORT[key])}'; ws['A1'].font = Font(bold=True, size=13, color=NAVYX)
        ws['A2'] = f'Source: {SOURCE} | Data as of {REPLBL}'; ws['A2'].font = Font(italic=True, color='64748B')
        shown = [x for x, on in ((('DATA', t.get('fact', '')), SHOW_DATA), (('SCREEN', t.get('act', '')), SHOW_SCREEN), (('METHOD', t.get('meth', '')), SHOW_METHOD)) if on]
        for r, (lab, txt) in enumerate(shown, 3):
            ws.cell(r, 1, f'{lab}: {txt}').alignment = WRAP; ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = max(30, 15 * (len(txt) // 150 + 1))
        ws.column_dimensions['A'].width = 34
        ws.page_setup.orientation = 'landscape'; ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True); ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        for j in range(2, 30): ws.column_dimensions[CL(j)].width = 15
        return ws
    def put(ws, b, r):
        df = b['df'].copy(); df.index.name = df.index.name or b.get('idx', 'Theme'); df = df.reset_index()
        if b.get('title'): ws.cell(r, 1, b['title']).font = Font(bold=True, color=NAVYX); r += 1
        for j, c in enumerate(df.columns, 1):
            cell = ws.cell(r, j, str(c)); cell.fill, cell.font, cell.alignment = HF, HFONT, Alignment(wrap_text=True, vertical='center')
        for i, row in enumerate(df.itertuples(index=False), 1):
            for j, v in enumerate(row, 1):
                cell = ws.cell(r + i, j, py(v)); f_ = b.get('fmts', {}).get(df.columns[j - 1])
                if f_: cell.number_format = f_
        b['pos'] = (r, r + 1, r + len(df), len(df.columns), list(df.columns)); return r, r + len(df)
    def ref(ws, b, col, hdr=False):
        h, f, l, n, cols = b['pos']; j = cols.index(col) + 1; return Reference(ws, min_col=j, min_row=h if hdr else f, max_row=l)
    def style_pts(series, colors, size=11, border='FFFFFF'):
        # per-point marker styling: custom fill colour, size and border stroke for each point
        series.dPt = []
        for i, colr in enumerate(colors):
            pt = DataPoint(idx=i); pt.marker = Marker(symbol='circle', size=size); pt.marker.graphicalProperties = GraphicalProperties(solidFill=colr)
            pt.marker.graphicalProperties.line.solidFill = border; series.dPt.append(pt)
    def axes(ch, xt=None, yt=None, yfmt=None, w=None, h=None, width=None, height=None):
        ch.x_axis.delete = False; ch.y_axis.delete = False; ch.height, ch.width = h or height or 9.5, w or width or 20
        ch.x_axis.tickLblPos = 'low'; ch.y_axis.tickLblPos = 'low'
        if xt: ch.x_axis.title = xt
        if yt: ch.y_axis.title = yt
        if yfmt: ch.y_axis.number_format = yfmt
    def chart_for(ws, b, c, slot=0):
        k = c['kind']; h, f, l, n, cols = b['pos']; ch = None
        if k in ('bar', 'col'):
            ch = BarChart(); ch.type = 'bar' if k == 'bar' else 'col'; ch.grouping = 'clustered'
            for v in c['vals']: ch.add_data(ref(ws, b, v, True), titles_from_data=True)
            ch.set_categories(ref(ws, b, cols[0])); axes(ch, yfmt=c.get('fmt'), w=c.get('w') or c.get('width'), h=c.get('h') or c.get('height'))
            if k == 'bar': ch.x_axis.scaling.orientation = 'maxMin'
            ch.x_axis.tickLblPos = 'low'; ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
            for a_ in ('showSerName', 'showCatName', 'showLegendKey', 'showPercent'): setattr(ch.dataLabels, a_, False)
            ch.dataLabels.numFmt = c.get('fmt', 'General')
        elif k == 'line':
            ch = LineChart()
            for v in c['vals']: ch.add_data(ref(ws, b, v, True), titles_from_data=True)
            ch.set_categories(ref(ws, b, cols[0])); axes(ch, yfmt=c.get('fmt'), w=c.get('w') or c.get('width'), h=c.get('h') or c.get('height'))
        elif k == 'dumbbell':
            ch = LineChart()
            for v in c['vals']: ch.add_data(ref(ws, b, v, True), titles_from_data=True)
            for s_, colr in zip(ch.series, (NAVYX, 'B42318')):
                s_.marker = Marker(symbol='circle', size=9); s_.marker.graphicalProperties = GraphicalProperties(solidFill=colr); s_.marker.graphicalProperties.line.solidFill = colr
                s_.graphicalProperties.line.noFill = True
            ch.hiLowLines = ChartLines(); ch.set_categories(ref(ws, b, cols[0])); axes(ch, None, None, c.get('fmt'), c.get('w'), c.get('h'))
            if c.get('log'): ch.y_axis.scaling.logBase = 10
        elif k == 'combo':
            bar_type = c.get('bar_type', 'bar' if c.get('orient', 'h') == 'h' else 'col')
            b_val = c.get('bar_val', c.get('val'))
            m_val = c.get('marker_val', c.get('val'))
            ch = BarChart(); ch.type = bar_type; ch.grouping = 'clustered'
            ch.add_data(ref(ws, b, b_val, True), titles_from_data=True)
            ch.set_categories(ref(ws, b, cols[0]))
            axes(ch, yfmt=c.get('fmt'), w=c.get('w') or c.get('width', 22), h=c.get('h') or c.get('height', 15))
            if ch.type == 'bar':
                ch.x_axis.scaling.orientation = 'maxMin'
                ch.x_axis.axPos = 'l'
                ch.y_axis.axPos = 'b'
            if c.get('hide_bars', True):
                ch.series[0].graphicalProperties.noFill = True
                ch.series[0].graphicalProperties.line.noFill = True
            else:
                ch.series[0].graphicalProperties.solidFill = c.get('bar_color', NAVYX)
            ch.legend = None
            ch2 = LineChart()
            ch2.add_data(ref(ws, b, m_val, True), titles_from_data=True)
            s2 = ch2.series[0]
            s2.graphicalProperties.line.noFill = True
            sz = c.get('size', 12)
            s2.marker.symbol = 'circle'; s2.marker.size = sz
            v_col = cols.index(m_val) + 1
            colors_list = c.get('colors')
            borders_list = c.get('borders')
            for idx, r in enumerate(range(f, l + 1)):
                v_val = ws.cell(r, v_col).value
                dp = DataPoint(idx=idx); dp.marker = Marker(symbol='circle', size=sz)
                if colors_list and idx < len(colors_list):
                    dp_color = colors_list[idx]
                else:
                    col_neg = c.get('marker_neg', 'B42318'); col_pos = c.get('marker_pos', '1A7F4B')
                    dp_color = col_neg if (v_val or 0) < 0 else col_pos
                dp.marker.graphicalProperties.solidFill = dp_color
                border_clr = borders_list[idx] if (borders_list and idx < len(borders_list)) else ('B42318' if (v_val or 0) < 0 else '1A7F4B')
                dp.marker.graphicalProperties.line.solidFill = border_clr
                dp.marker.graphicalProperties.line.width = 15000
                s2.dPt.append(dp)
            s2.dLbls = DataLabelList(); s2.dLbls.showVal = True; s2.dLbls.numFmt = c.get('fmt', '0.0%')
            s2.dLbls.dLblPos = 'bestFit'
            for a_ in ('showSerName', 'showCatName', 'showLegendKey', 'showPercent'): setattr(s2.dLbls, a_, False)
            ch += ch2
        elif k == 'bubble':
            ch = BubbleChart(); ch.style = 18; ch.bubbleScale = 50; ch.legend = None
            for r in range(f, l + 1):
                s_ = Series(values=Reference(ws, min_col=cols.index(c['y']) + 1, min_row=r), xvalues=Reference(ws, min_col=cols.index(c['x']) + 1, min_row=r),
                            zvalues=Reference(ws, min_col=cols.index(c['z']) + 1, min_row=r))
                s_.tx = SeriesLabel(strRef=StrRef(f"'{ws.title}'!$A${r}")); s_.dLbls = DataLabelList(); s_.dLbls.showSerName = True
                for a_ in ('showVal', 'showCatName', 'showLegendKey', 'showPercent'): setattr(s_.dLbls, a_, False)
                ch.series.append(s_)
            axes(ch, c.get('xt'), c.get('yt'), c.get('fmt'), c.get('w'), c.get('h')); ch.x_axis.number_format = c.get('fmt', 'General')
        elif k == 'scatter':
            ch = ScatterChart(); ch.style = 13; ch.scatterStyle = 'lineMarker'; ch.legend = None
            s_ = Series(ref(ws, b, c['y']), ref(ws, b, c['x']), title='Stocks'); s_.marker = Marker(symbol='circle', size=5); s_.graphicalProperties.line.noFill = True; ch.series.append(s_)
            t_ = c.get('thr')
            if t_ is not None:
                th_, tf, tl, tn, tc = t_['pos']
                for r0 in (tf, tf + 2):
                    ts = Series(Reference(ws, min_col=2, min_row=r0, max_row=r0 + 1), Reference(ws, min_col=1, min_row=r0, max_row=r0 + 1), title='Threshold')
                    ts.marker = Marker(symbol='none'); ts.graphicalProperties.line.solidFill = 'B42318'; ts.graphicalProperties.line.dashStyle = 'dash'; ch.series.append(ts)
            axes(ch, c.get('xt'), c.get('yt')); ch.x_axis.scaling.logBase = 10; ch.y_axis.scaling.logBase = 10
        if ch is not None:
            ch.title = c.get('title'); right = n <= 8
            ws.add_chart(ch, f'{CL(n + 2)}{h + 21 * slot}' if right else f'A{l + 2 + 24 * slot}'); return max(l + 3, h + 21 * (slot + 1)) if right else l + 24 * (slot + 1)
        return l + 3
    def cf(ws, b):
        h, f, l, n, cols = b['pos']
        for names_, kind in b.get('cf', []):
            js = [cols.index(x) + 1 for x in names_ if x in cols]
            if not js: continue
            rng = f'{CL(min(js))}{f}:{CL(max(js))}{l}'
            ws.conditional_formatting.add(rng, ColorScaleRule(start_type='min', start_color='F8696B', mid_type='num', mid_value=0, mid_color='FFFFFF', end_type='max', end_color='63BE7B') if kind == 'div'
                                          else ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='max', end_color=NAVYX))
    def build(key, blocks):
        ws = mk_sheet(key); r = 7
        for b in blocks:
            r0, last = put(ws, b, r); r = last + 3
        r = 7
        for b in blocks:
            cf(ws, b); h, f, l, n, cols = b['pos']
            specs = b.get('charts') or ([b['chart']] if b.get('chart') else [])
            for slot, c_ in enumerate(specs): r = max(r, chart_for(ws, b, c_, slot))
            r = max(r, l + 3)
            r = max(r, l + 3)
        for b in blocks:
            for j, c in enumerate(b['pos'][4], 1):
                ws.column_dimensions[CL(j)].width = max(ws.column_dimensions[CL(j)].width or 0, min(22, max(12, len(str(c)) * 0.9)))
    def D(**kw): return kw
    def cleaned(df): return df.replace([np.inf, -np.inf], np.nan)
    builders = {}
    def peers_():
        d = peer.loc[pr, ['n', 'aum', 'flow_q', 'flow_ytd', 'flow_1y', 'ogr_q', 'ogr_ytd', 'ogr_1y']]
        d.columns = ['ETFs', 'AuM ($mn)', 'QTD absolute flow ($mn)', 'YTD absolute flow ($mn)', '1Y absolute flow ($mn)', 'QTD normalized flow', 'YTD normalized flow', '1Y normalized flow']
        sh = pd.DataFrame({'Share': [(st.peer == THEME_VALUE).mean(), st[st.peer == THEME_VALUE].aum.sum() / max(st.aum.sum(), 1e-9)]}, index=['Themes share of equity ETF listings', 'Themes share of equity ETF AuM'])
        fm = {c: MNF for c in d.columns if 'mn' in c}; fm.update({c: PCT for c in d.columns if 'normalized' in c})
        return [D(df=d, idx='Peer group', fmts=fm, chart=D(kind='bar', vals=['QTD normalized flow', 'YTD normalized flow', '1Y normalized flow'], fmt=PCT, title='Normalized flow (% of base AuM)')), D(df=sh, idx='Measure', fmts={'Share': PCT})]
    def pol_():
        d = pol[['flow_q', 'ogr_q', 'aum_q0']]; d.columns = ['QTD absolute flow ($mn)', 'QTD normalized flow', 'Base AuM ($mn)']
        return [D(df=d, fmts={'QTD absolute flow ($mn)': MNF, 'QTD normalized flow': PCT, 'Base AuM ($mn)': MNF}, chart=D(kind='bar', vals=['QTD absolute flow ($mn)'], fmt=INTF, title='QTD absolute flow ($mn)'))]
    def card_(df):
        m1, m2 = qm[0].strftime('%b'), qm[-1].strftime('%b'); c = ['n', 'aum', 'flow_q', 'ogr_q', 'flow_ytd', 'ogr_ytd', 'm_first', 'm_last', 'ret_q', 'ret_ytd', 'ret_1y', 'ex_acwi_1y', 'neff', 'launch_q']
        c = [x for x in c if x in th.columns]; d = th.loc[df.index, c]
        nm = dict(n='ETFs', aum='AuM ($mn)', flow_q='QTD absolute flow ($mn)', ogr_q='QTD normalized flow', flow_ytd='YTD absolute flow ($mn)', ogr_ytd='YTD normalized flow', m_first=f'{m1} flow ($mn)', m_last=f'{m2} flow ($mn)',
                  ret_q='QTD return', ret_ytd='YTD return', ret_1y='1Y return', ex_acwi_1y='1Y excess vs MSCI ACWI', neff='N.eff', launch_q='QTD launches'); d = d.rename(columns=nm)
        fm = {v: (PCT if ('normalized' in v or 'return' in v or 'excess' in v) else (MNF if '$mn' in v else (DEC1 if v == 'N.eff' else INTF))) for v in d.columns}
        return [D(df=d, fmts=fm)]
    def trend_():
        d = cum.loc[sel].T; d.index = [c.strftime('%b') for c in d.index]; return [D(df=d, idx='Month', fmts={c: MNF for c in d.columns}, chart=D(kind='line', vals=list(d.columns), fmt=INTF, title='Cumulative YTD absolute flows ($mn)'))]
    def perf_():
        d = pd.DataFrame(hm_, index=hr_, columns=cols_); return [D(df=d, fmts={c: PCT for c in d.columns}, cf=[(['QTD', 'YTD', '1Y'], 'div'), ([c for c in cols_ if c.startswith('1Y vs')], 'div')])]
    def asym_():
        d = th[th.ret_1y.notna() & th.dd.notna()][['aum', 'ret_1y', 'dd', 'vol', 'sharpe']].copy(); d['Return ÷ |drawdown|'] = d.ret_1y / d.dd.abs(); d.columns = ['AuM ($mn)', '1Y return', 'Max drawdown', '1Y volatility', 'Sharpe', 'Return ÷ |drawdown|']
        return [D(df=d, fmts={'AuM ($mn)': MNF, '1Y return': PCT, 'Max drawdown': PCT, '1Y volatility': PCT, 'Sharpe': DEC2, 'Return ÷ |drawdown|': DEC2}, chart=D(kind='bubble', x='Max drawdown', y='1Y return', z='AuM ($mn)', xt='1Y max drawdown', yt='1Y total return', fmt=PCT, title='Asymmetry map (bubble = AuM)'))]
    def chase_():
        d = fcr.sort_values('gap').set_index('theme')[['tw', 'dw', 'gap']]; d.columns = ['Time-weighted 1Y return', 'Dollar-weighted 1Y return', 'Gap']
        gaps = [float(x) for x in d['Gap']]
        g_min, g_max = min(gaps + [0]), max(gaps + [0])
        def _c(g):
            if g < 0:
                t = g / g_min if g_min < 0 else 0
                rgb = (int(round(255 + (248 - 255) * t)), int(round(255 + (105 - 255) * t)), int(round(255 + (107 - 255) * t)))
            else:
                t = g / g_max if g_max > 0 else 0
                rgb = (int(round(255 + (99 - 255) * t)), int(round(255 + (190 - 255) * t)), int(round(255 + (123 - 255) * t)))
            return f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'
        bubble_colors = [_c(g) for g in gaps]
        border_colors = ['B42318' if g < 0 else '1A7F4B' for g in gaps]
        ch = D(kind='combo', bar_type='bar', bar_val='Gap', hide_bars=True, marker_val='Gap',
               colors=bubble_colors, borders=border_colors, fmt=PCT, height=15, width=22,
               title='The Flow-Chasing Penalty: Investor Return Gap by Theme (Spheres: Red=Penalty, Green=Premium)')
        return [D(df=d, fmts={c: PCT for c in d.columns}, cf=[(['Gap'], 'div')], chart=ch)]
    def launch_():
        c1_ = pd.DataFrame({'Active': [int((stT.new_q & stT.active & (stT[M_DOM] == d)).sum()) for d in doms], 'Passive': [int((stT.new_q & ~stT.active & (stT[M_DOM] == d)).sum()) for d in doms]}, index=doms)
        c2_ = pd.DataFrame({d: [ashare(f_, d) for f_ in fns] for d in doms}, index=labs)
        nl_ = stT[stT.new_q].sort_values('aum', ascending=False); l_ = nl_[[M_NAME, 'theme', M_STYLE, M_DOM, M_TER, 'aum']].copy(); l_.index = nl_[M_TICKER]; l_.columns = ['Fund', LABEL_T2, 'Style', 'Domicile', 'TER (%)', 'AuM ($mn)']
        return [D(df=c1_, title='QTD launches by domicile', idx='Domicile', fmts={}, chart=D(kind='bar', vals=['Active', 'Passive'], fmt='0', title='QTD launches')),
                D(df=c2_, title='Active share by domicile', idx='Measure', fmts={d: PCT for d in doms}, chart=D(kind='bar', vals=list(doms), fmt='0%', title='Active share')),
                D(df=l_, title='QTD launches (all)', idx='Ticker', fmts={'TER (%)': DEC2, 'AuM ($mn)': MNF})]
    def conc_():
        cs = conc.sort_values('hhi', ascending=False); top = cs.head(12); rows = [(r.label, r.N, r.neff) for r in top.itertuples()] + [(f'{k}', v[0], 1e4 / v[1]) for k, v in refs.items()]
        c1_ = pd.DataFrame(rows, columns=['Index', 'Nominal constituents N', 'Effective constituents N.eff']).set_index('Index')
        full = cs.set_index('label')[['etf', 'theme', 'aum', 'n_etf', 'N', 'hhi', 'neff']]; full['N.eff ÷ N'] = full.neff / full.N; full.columns = ['Representative ETF (ISIN)', LABEL_T2, 'AuM tracking ($mn)', 'ETFs tracking', 'N', 'HHI', 'N.eff', 'N.eff ÷ N']
        return [D(df=c1_, title='Chart data: 12 most concentrated thematic indices and reference indices', idx='Index', fmts={c: DEC1 for c in c1_.columns}, chart=D(kind='dumbbell', vals=list(c1_.columns[:2]), fmt='0', log=True, title='N vs. N.eff (log scale)')),
                D(df=full, title='All thematic indices', idx='Index', fmts={'AuM tracking ($mn)': MNF, 'N': INTF, 'HHI': INTF, 'N.eff': DEC1, 'N.eff ÷ N': PCT})]
    def own_():
        d = g.sort_values('float', ascending=False)[['name', 'theme', 'n_etfs', 'pos', 'mcap', 'float', 'dtl', 'danger']].copy(); d.columns = ['Stock', LABEL_T2, 'ETFs holding', 'ETF position ($mn)', 'Market cap ($mn)', '% of market cap held', 'Days to liquidate', f'In {LBL_STRESS}']
        d['% of market cap held'] = d['% of market cap held'] / 100
        xs, ys = d['Days to liquidate'], d['% of market cap held'] * 100
        thr = pd.DataFrame({'x': [DANGER_DTL, DANGER_DTL, float(xs.min()) * .8, float(xs.max()) * 1.2], 'y': [float(ys.min()) * .8, float(ys.max()) * 1.2, DANGER_FLOAT, DANGER_FLOAT]}); thr.index = ['Days threshold', 'Days threshold', 'Share threshold', 'Share threshold']
        d2 = d.copy(); d2['% held (x100, chart)'] = ys
        bl = D(df=d2, title='All stocks (sorted by % of market cap held)', idx='ISIN', fmts={'ETF position ($mn)': MNF, 'Market cap ($mn)': INTF, '% of market cap held': PCT, 'Days to liquidate': DEC1, '% held (x100, chart)': DEC2})
        tb = D(df=thr.rename(columns={'x': 'Days to liquidate', 'y': '% of market cap held'}), title='Threshold lines for the chart', idx='Line', fmts={})
        bl['chart'] = D(kind='scatter', x='Days to liquidate', y='% held (x100, chart)', xt='Days to liquidate (log)', yt='% of market cap held (log)', thr=tb, title=f'{LBL_STRESS}: ownership vs. days to liquidate')
        return [bl, tb]
    def ovl_():
        d = pd.DataFrame(O, index=lab, columns=[st.at[reps[t], M_TICKER] for t in ks]); return [D(df=d, idx='ETF · theme', fmts={c: DEC1 for c in d.columns}, cf=[(list(d.columns), 'seq')])]
    def corr_():
        d = C.copy(); return [D(df=d, idx='Theme', fmts={c: DEC2 for c in d.columns}, cf=[(list(d.columns), 'seq')])]
    def lg_():
        d = th.sort_values('aum', ascending=False)[['n', 'aum', 'flow_q', 'ogr_q', 'ogr_ytd', 'ret_q', 'ret_ytd', 'ret_1y', 'vol', 'sharpe', 'dd', 'hhi', 'neff']].copy()
        d.columns = ['ETFs', 'AuM ($mn)', 'QTD absolute flow ($mn)', 'QTD normalized flow', 'YTD normalized flow', 'QTD return', 'YTD return', '1Y return', '1Y volatility', 'Sharpe', 'Max drawdown', 'HHI', 'N.eff']
        fm = {c: PCT for c in d.columns if ('normalized' in c or 'return' in c or c in ('1Y volatility', 'Max drawdown'))}; fm.update({'AuM ($mn)': MNF, 'QTD absolute flow ($mn)': MNF, 'Sharpe': DEC2, 'HHI': INTF, 'N.eff': DEC1})
        return [D(df=d, fmts=fm)]
    builders = dict(peers=peers_, pol=pol_, gain=lambda: card_(gain), lose=lambda: card_(lose), trend=trend_, perf=perf_, asym=asym_, chase=chase_, launch=launch_, conc=conc_, own=own_, ovl=ovl_, corr=corr_, lg=lg_)
    idx = wb.create_sheet('Index', 0); idx['A1'] = f'Thematic ETF Monitor & Strategy Chartbook: data workbook, {Q_LBL}'; idx['A1'].font = Font(bold=True, size=14, color=NAVYX)
    idx['A2'] = f'Source: {SOURCE} | Data as of {REPLBL}. Each sheet holds the data behind one exhibit and native Excel charts or color scales built from those cells, so everything can be reformatted.'; idx['A2'].font = Font(italic=True, color='64748B')
    for j, h_ in enumerate(('Exhibit', 'Section', 'Title', 'Sheet'), 1): c_ = idx.cell(4, j, h_); c_.fill, c_.font = HF, HFONT
    rr = 5
    for key in sorted(NUMS, key=NUMS.get):
        try:
            blocks = builders[key](); build(key, blocks); status = wb[f'Ex{NUMS[key]} {SHORT[key]}'[:31]].title
        except Exception as e:
            log(f'Excel sheet for exhibit {NUMS[key]} ({key}) skipped: {type(e).__name__}: {str(e)[:100]}', 'WARN'); status = 'not available'
        idx.cell(rr, 1, NUMS[key]); idx.cell(rr, 2, SECS[key]); idx.cell(rr, 3, TXT.get(key, {}).get('title', key)); idx.cell(rr, 4, status); rr += 1
    for col_, w in zip('ABCD', (10, 48, 90, 26)): idx.column_dimensions[col_].width = w
    wb.save(path); log(f'Excel workbook written: {path}')
if OUT_XLSX and MAKE_EXCEL:
    try:
        out_path = os.path.join(OUT_DIR, OUT_XLSX)
        write_excel(out_path); print('Excel written:', out_path)
        _here = os.path.dirname(os.path.abspath(__file__))
        for sync_dir in [os.path.abspath(os.path.join(_here, '..', '..')), os.path.abspath(os.path.join(_here, '..'))]:
            if os.path.isdir(sync_dir):
                try: shutil.copy2(out_path, os.path.join(sync_dir, OUT_XLSX))
                except Exception: pass
    except Exception as e: log(f'Excel export failed: {type(e).__name__}: {e}', 'WARN')

def to_pdf(src, dst):
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(); pg = b.new_page(); pg.goto('file://' + os.path.abspath(src)); pg.pdf(path=dst, format='A4', print_background=True, margin=dict(top='10mm', bottom='10mm', left='8mm', right='8mm')); b.close()
        return 'playwright'
    except Exception as e: log(f'Playwright PDF failed: {str(e)[:120]}', 'WARN')
    if shutil.which('wkhtmltopdf'):
        subprocess.run(['wkhtmltopdf', '--enable-local-file-access', '-q', src, dst]); return 'wkhtmltopdf'
    return None
if MAKE_PDF:
    how = to_pdf(hp, os.path.join(OUT_DIR, OUT_PDF)); print('PDF written via', how) if how else print('PDF NOT created: install playwright (pip install playwright; playwright install chromium) or wkhtmltopdf')
    if not MAKE_HTML and os.path.exists(hp): os.remove(hp)
else: print('PDF skipped (MAKE_PDF = False)')
