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
x['win_ok'] = weq(X_QTD_S, PQE) & weq(X_QTD_E, REP) & weq(X_YTD_S, YS) & weq(X_YTD_E, REP) & weq(X_1Y_S, Y1) & weq(X_1Y_E, REP)