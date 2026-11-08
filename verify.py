# -*- coding: utf-8 -*-
import pandas as pd, numpy as np, re, json
SRC = r'c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_refreshed.xlsx'
xl = pd.ExcelFile(SRC)
def kn(c): return re.sub(r'（[^（）]*）','',c).strip()

# 1) 半钢产能利用率 异常核查
df = pd.read_excel(xl,'下游'); cols=list(df.columns); cols[0]='dt'; df.columns=cols
df['dt']=pd.to_datetime(df['dt'],errors='coerce'); df=df.dropna(subset=['dt'])
col='半钢轮胎：产能利用率：中国（周）'
s=df[['dt',col]].dropna(); s['dt']=pd.to_datetime(s['dt']); s=s.set_index('dt')[col].sort_index()
print('半钢 util: n=%d  latest=%.3f @ %s  min=%.3f max=%.3f'%(len(s),s.iloc[-1],s.index.max().date(),s.min(),s.max()))
big=s[s>120]
print('>120 的点数:',len(big))
for d,v in big.items(): print('   ',d.date(),'%.2f'%v)
# 2015-06-30 相邻
for d in [pd.Timestamp('2015-06-28'),pd.Timestamp('2015-06-29'),pd.Timestamp('2015-06-30'),pd.Timestamp('2015-07-01'),pd.Timestamp('2015-07-02')]:
    v=s.asof(d); print('   asof',d.date(),'->',(None if pd.isna(v) else round(float(v),3)))

# 2) 泰混通关-国内市场价 核査
dfm=pd.read_excel(xl,'中游'); colsm=list(dfm.columns); colsm[0]='dt'; dfm.columns=colsm
dfm['dt']=pd.to_datetime(dfm['dt'],errors='coerce'); dfm=dfm.dropna(subset=['dt'])
c='泰混通关-国内市场价'
ss=dfm[['dt',c]].dropna(); ss['dt']=pd.to_datetime(ss['dt']); ss=ss.set_index('dt')[c].sort_index()
print('\n泰混通关-国内市场价: n=%d latest=%.2f @ %s min=%.2f max=%.2f'%(len(ss),ss.iloc[-1],ss.index.max().date(),ss.min(),ss.max()))

# 3) 全钢/半钢利润 单位与量纲
for c in ['全钢轮胎利润（元/吨）','半钢轮胎利润（元/吨）']:
    t=df[['dt',c]].dropna(); t['dt']=pd.to_datetime(t['dt']); t=t.set_index('dt')[c].sort_index()
    print(c,'n=%d latest=%.2f @ %s min=%.2f max=%.2f'%(len(t),t.iloc[-1],t.index.max().date(),t.min(),t.max()))

# 4) 各段最新日期（日 vs 周）
print('\n=== 各指标截至日期 ===')
for sh in ['上游','中游','下游']:
    d=pd.read_excel(xl,sh); cc=list(d.columns); cc[0]='dt'; d.columns=cc
    d['dt']=pd.to_datetime(d['dt'],errors='coerce'); d=d.dropna(subset=['dt'])
    for c in [x for x in d.columns if x!='dt']:
        sub=d[['dt',c]].dropna()
        if sub.empty: continue
        print('  %s | %s | asof=%s n=%d'%(sh, kn(c), sub['dt'].max().date(), len(sub)))
