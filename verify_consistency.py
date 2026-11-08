# -*- coding: utf-8 -*-
import pandas as pd, numpy as np, json, re
SRC = r'c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_refreshed.xlsx'
P = json.load(open(r'c:\Users\Lenovo\WorkBuddy\20261021112446\site_data.json', encoding='utf-8'))
xl = pd.ExcelFile(SRC)
def kn(c): return re.sub(r'（[^（）]*）','',c).strip()
# 重新读 Excel 真值
ex = {}
for sh in ['上游','中游','下游']:
    d = pd.read_excel(xl, sh); cc=list(d.columns); cc[0]='dt'; d.columns=cc
    d['dt']=pd.to_datetime(d['dt'],errors='coerce'); d=d.dropna(subset=['dt'])
    for c in [x for x in d.columns if x!='dt']:
        sub=d[['dt',c]].dropna()
        if sub.empty: continue
        s=sub.set_index('dt')[c].sort_index()
        ex[(sh,kn(c))]=s

rows=[]
mismatch=0
for sh,seg in P['segments'].items():
    for ind in seg['indicators']:
        key=(sh,ind['name'])
        s=ex.get(key)
        if s is None:
            rows.append((ind['name'],'EXCEL缺失','-','-','-')); mismatch+=1; continue
        # 半钢排除 7 个异常点
        if '产能利用率' in ind['name']:
            s=s[(s<=100)&(s>=0)]
        latest_date=s.index.max(); latest=float(s.loc[latest_date])
        n=len(s)
        st=ind['stats']
        ok_v = abs(round(latest,4)-st['latest'])<1e-3
        ok_d = str(latest_date.date())==st['latest_date']
        ok_n = n==st['n']
        flag = '' if (ok_v and ok_d and ok_n) else '  <-- 不一致'
        if flag: mismatch+=1
        rows.append((ind['name'], '%.4f'%latest, st['latest'], str(latest_date.date()), st['latest_date'], n, st['n'], flag))

print('=== Excel ↔ JSON 逐项一致性（最新值/日期/样本量）===')
print('%-34s %12s %12s %12s %12s %6s %6s' % ('name','excel_val','json_val','excel_date','json_date','ex_n','jn'))
for r in rows:
    if len(r)==8:
        print('%-34s %12s %12s %12s %12s %6s %6s%s' % r)
    else:
        print('%-34s %s' % (r[0], r[1]))
print('\n不一致项数:', mismatch, '/', len(rows))

# 半钢异常专项
ind=[x for x in P['segments']['下游']['indicators'] if x['name']=='半钢轮胎：产能利用率：中国'][0]
print('\n半钢 util JSON: latest',ind['stats']['latest'],'@',ind['stats']['latest_date'],'n',ind['stats']['n'],'excluded_n',ind['excluded_n'],'hist_q',ind['stats']['hist_q'],'mean',ind['stats']['mean'])
se=r"c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_refreshed.xlsx"
s=ex[('下游','半钢轮胎：产能利用率：中国')]
print('半钢 Excel 原始 n=',len(s),' 排除后 n=',len(s[(s<=100)&(s>=0)]),' 最大值(原始)=',s.max())

# 核心最新值表
print('\n=== 核心指标最新值与截至日期 ===')
for d in P['summary']['core']:
    print('  %-22s 最新 %.3f %s | asof %s | hist_q %.1f | 分位变化 %s | %s' % (
        d['name'], d['latest'], d['unit'],
        [x for x in P['segments'][d['sheet']]['indicators'] if x['name']==d['name']][0]['stats']['latest_date'],
        d['hist_q'], d['chg20_q'], '理论/待确认' if d['unconfirmed'] else ''))

# 刷新报告
print('\n=== 刷新状态 ===')
print(json.dumps(P['meta']['refresh'] if 'refresh' in P['meta'] else P['summary']['refresh'], ensure_ascii=False, indent=1))

# oct window
print('\n=== 10月窗口对比 ===')
for o in P['summary']['oct_window']:
    print('  %-30s 2026=%.2f 往年均值=%.2f delta=%+.2f %s' % (o['name'],o['cur'],o['prev_avg'],o['delta'],o['flag']))
print('  注:', P['summary']['oct_note'])
