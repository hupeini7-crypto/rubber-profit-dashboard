import json, sys

sys.stdout.reconfigure(encoding="utf-8")
d = json.load(open("site_data.json", encoding="utf-8"))
seg = d["segments"]
inds = []
for k in ("上游", "中游", "下游"):
    inds.extend(seg[k]["indicators"])

for ind in inds:
    if ind.get("is_core"):
        s = ind["stats"]
        print(f"{ind['name']:<16} {s['latest']:>10} {ind['unit']:<10} asof {s['latest_date']} | q={s['hist_q']:>5} | dq20={s.get('chg20_q'):>6} | 20d={s.get('chg20')}")
