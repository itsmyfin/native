log(f'Universe: {len(st)} equity ETF share classes (of {len(m)} rows; ETC/ETN and non-equity excluded)')
st[M_TICKER] = st[M_TICKER].astype(str).str.strip()
st[M_STRAT] = st[M_STRAT].astype(str).str.strip()
st.loc[st[M_STRAT].str.lower().isin(['themes', 'thematic', 'thematics', 'theme']), M_STRAT] = THEME_VALUE
st['peer'] = st[M_STRAT].where(st[M_STRAT].isin(PEER_GROUPS), 'Other')



x = x.drop_duplicates(X_TICKER).set_index(X_TICKER)
x.index = x.index.astype(str).str.strip()
for c in (X_QTD, X_YTD, X_1Y, X_VOL, X_SHARPE, X_DD): x[c] = pd.to_numeric(x[c], errors='coerce')

def to_dt(s):
    if pd.api.types.is_datetime64_any_dtype(s): return s.dt.normalize()
    num = pd.to_numeric(s, errors='coerce')
    is_serial = num.notna() & (num > 30000) & (num < 60000)
    res = pd.Series(pd.NaT, index=s.index, dtype='datetime64[ns]')
    if is_serial.any(): res.loc[is_serial] = pd.to_datetime(num[is_serial], unit='D', origin='1899-12-30').dt.normalize()
    rem = ~is_serial & s.notna()
    if rem.any():
        s_str = s[rem].astype(str).str.strip().replace({'nan': None, 'None': None, 'NaT': None, '': None})
        is_8d = s_str.str.match(r'^\d{8}$')
        if is_8d.any():
            i8 = s_str[is_8d].index; res.loc[i8] = pd.to_datetime(s_str.loc[i8], format='%Y%m%d', errors='coerce').dt.normalize(); rem = rem & ~is_8d
        if rem.any(): res.loc[rem] = pd.to_datetime(s_str.loc[rem], errors='coerce', format='mixed').dt.normalize()
    return res

for c in (X_QTD_S, X_QTD_E, X_YTD_S, X_YTD_E, X_1Y_S, X_1Y_E): x[c] = to_dt(x[c])
for c in (X_QTD, X_YTD, X_1Y, X_VOL):
    if x[c].abs().quantile(.99) > 5: log(f'{c} looks like percent units; divided by 100', 'WARN'); x[c] = x[c] / 100
if x[X_DD].median() > 0: log('Max drawdown was positive; sign flipped', 'WARN'); x[X_DD] = -x[X_DD]

def weq(col, d): dt = x[col]; return dt.notna() & ((dt - d).abs().dt.days <= 3)
x['win_ok'] = weq(X_QTD_S, PQE) & weq(X_QTD_E, REP) & weq(X_YTD_S, YS) & weq(X_YTD_E, REP) & (x[X_1Y_S].isna() | weq(X_1Y_S, Y1)) & (x[X_1Y_E].isna() | weq(X_1Y_E, REP))
bad = list(x.index[~x.win_ok])
if bad: log(f'{len(bad)} metrics rows fail the date-window assertion and are excluded from return/risk stats: {names(bad, 8)}', 'WARN')
xm = x.rename(columns={X_QTD: 'ret_q', X_YTD: 'ret_ytd', X_1Y: 'ret_1y', X_VOL: 'vol', X_SHARPE: 'sharpe', X_DD: 'dd'})[['ret_q', 'ret_ytd', 'ret_1y', 'vol', 'sharpe', 'dd', 'win_ok']]
st = st.join(xm, on=M_TICKER); st['ret_ok'] = st.win_ok.eq(True)

if not (st.ret_ok & (st.peer == THEME_VALUE)).any():
    thm_total = int((st.peer == THEME_VALUE).sum())
    metrics_match = int(st[M_TICKER].isin(x.index).sum())
    diag = []
    if thm_total == 0:
        diag.append(f"Zero ETFs in Main file categorized under peer '{THEME_VALUE}' (Equity Strategy 1 values found: {list(st[M_STRAT].unique())})")
    if metrics_match == 0:
        diag.append(f"Zero tickers in Main file matched ETF_metrics.xlsx (Sample Main: {list(st[M_TICKER])[:3]}, Sample Metrics: {list(x.index)[:3]})")
    file_ends = [str(d.date()) for d in x[X_QTD_E].dropna().unique()[:3]]
    if file_ends and not weq(X_QTD_E, REP).any():
        diag.append(f"Date mismatch: REPORT_DATE is '{REP.date()}', but ETF_metrics.xlsx dates end on {file_ends}. Update REPORT_DATE in build_chartbook.py (Line 53) to match.")
    diag_str = '; '.join(diag) if diag else f"Dates in ETF_metrics.xlsx do not match QTD [{PQE.date()} to {REP.date()}], YTD [{YS.date()} to {REP.date()}], 1Y [{Y1.date()} to {REP.date()}]. Actual sample dates found: QTD_Start={list(x[X_QTD_S].dropna().unique()[:2])}, QTD_End={file_ends}"
    fail(f'No thematic ETF passed the metrics date-window assertion. {diag_str}')