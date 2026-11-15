import re, json, sys

sys.stdout.reconfigure(encoding="utf-8")
html = open(r"c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_年度季节性图表.html", encoding="utf-8").read()

for k in ["renderMonthWindow", "renderCorrMatrix", "lag-sel", "corr-heatmap", "setCorrMode", "month_window", "lag_scan"]:
    print(f"{k}: {html.count(k)}")
for k in ["renderOctWindow", "oct_window", "国庆(10月)窗口"]:
    print(f"[旧残留] {k}: {html.count(k)}")

m = re.search(r"const P = (\{.*?\});\n", html, re.S)
d = json.loads(m.group(1))
print("JSON ok | divergence:", len(d["summary"]["market_divergence"]))

cm = d["summary"]["corr_matrix"]
print("matrix weekly labels:", len(cm["weekly"]["labels"]), "| daily:", len(cm["daily"]["labels"]))
print("weekly sample row:", cm["weekly"]["labels"][:4], cm["weekly"]["r"][0][:4])
print("window meta:", d["summary"]["month_window_meta"])

l0 = d["linkage"][0]
m0 = l0["modes"]["diff"]["lag_scan"]
print("lag_scan freqs:", list(m0.keys()))
for f, sc in m0.items():
    print(f"  {f}: lags 0..{sc['lags'][-1]} best={sc['best_lag']} r={sc['best_corr']} n={sc['n']}")

js = html[html.rindex("<script>")+8:html.rindex("</script>")]
for a, b in [("{","}"),("(",")"),("[","]")]:
    print("pair", a+b, "balance", js.count(a)-js.count(b))
