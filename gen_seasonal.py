import pandas as pd
import json

SRC = r'D:\AAA东方嘉信\备份数据\上下游利润.xlsx'
OUT = r'c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_年度季节性图表.html'

# 用户指定的 sheet 顺序
SHEET_ORDER = ['上游', '中游', '下游']
SHEET_TITLE = {'上游': '上游 · 原料与割胶利润', '中游': '中游 · 贸易与盘面利润', '下游': '下游 · 轮胎与成品利润'}

xl = pd.ExcelFile(SRC)
available = xl.sheet_names
sheets_data = {}

for s in SHEET_ORDER:
    if s not in available:
        continue
    df = pd.read_excel(xl, s)
    # 第一列均为日期列，名称可能为 NaT/空，统一改为“时间”
    cols = list(df.columns)
    cols[0] = '时间'
    df.columns = cols
    df['时间'] = pd.to_datetime(df['时间'], errors='coerce')
    df = df.dropna(subset=['时间']).copy()
    df['年'] = df['时间'].dt.year
    df['月'] = df['时间'].dt.month

    indicators = [c for c in df.columns if c not in ('时间', '年', '月')]
    charts = []
    for col in indicators:
        sub = df[['年', '月', col]].dropna(subset=[col])
        if sub.empty:
            continue
        pivot = sub.pivot_table(index='月', columns='年', values=col, aggfunc='mean')
        pivot = pivot.reindex(range(1, 13))
        years = [int(y) for y in pivot.columns]
        # 多年均值（跨年逐月平均）
        seasonal = [None if pd.isna(v) else round(float(v), 4) for v in pivot.mean(axis=1).values]
        series = []
        for y in years:
            yvals = pivot[y]
            series.append({
                'year': y,
                'data': [None if pd.isna(v) else round(float(v), 4) for v in yvals.values]
            })
        charts.append({'name': str(col), 'years': years, 'series': series, 'seasonal': seasonal})
    sheets_data[s] = charts

payload = {'sheets': sheets_data, 'titles': SHEET_TITLE}
data_json = json.dumps(payload, ensure_ascii=False)

html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>上下游利润 · 年度季节性图表</title>
<style>
  * { box-sizing: border-box; }
  body { margin:0; font-family: "Microsoft YaHei", Arial, sans-serif; background:#f4f6f9; color:#1f2937; }
  header { background: linear-gradient(135deg,#0f4c81,#1f7a8c); color:#fff; padding:24px 32px; }
  header h1 { margin:0; font-size:24px; font-weight:700; }
  header p { margin:6px 0 0; opacity:.85; font-size:13px; }
  .sheet { margin:24px 32px; }
  .sheet h2 { font-size:19px; border-left:5px solid #1f7a8c; padding-left:12px; margin:0 0 16px; color:#0f4c81; }
  .grid { display:grid; grid-template-columns: repeat(auto-fill, minmax(440px,1fr)); gap:18px; }
  .card { background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,.06); padding:14px 16px 6px; }
  .card h3 { margin:0 0 4px; font-size:14px; font-weight:600; color:#334155; }
  .card .hint { font-size:11px; color:#94a3b8; margin-bottom:6px; }
  .chart { width:100%; height:300px; }
  footer { text-align:center; color:#94a3b8; font-size:12px; padding:24px; }
  .legend-note { font-size:12px; color:#64748b; margin:0 0 8px; }
  .legend-note b { color:#0f4c81; }
</style>
</head>
<body>
<header>
  <h1>上下游利润 · 年度季节性图表</h1>
  <p>数据来源：上下游利润.xlsx　|　口径：日度数据 → 逐月均值　|　X 轴为 1–12 月，细线为各年度，加粗线为多年均值（季节性规律）</p>
</header>
<div id="app"></div>
<footer>由 WorkBuddy 自动生成</footer>
<script src="echarts.min.js"></script>
<script>
const PAYLOAD = __DATA__;
const MONTHS = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
const app = document.getElementById('app');
const palette = ['#9aa7b1','#b6c2cc','#cdd6dd','#7f8c99','#aeb9c4','#6b7785','#c0cad2','#8a97a3','#d3dae0','#5d6b78','#b0bcc6','#94a3b1'];

Object.keys(PAYLOAD.sheets).forEach(sheet => {
  const wrap = document.createElement('div');
  wrap.className = 'sheet';
  const h2 = document.createElement('h2');
  h2.textContent = PAYLOAD.titles[sheet] || sheet;
  wrap.appendChild(h2);
  const note = document.createElement('p');
  note.className = 'legend-note';
  note.innerHTML = '<b>读图：</b>每条细线是一个自然年的月度走势；<b>加粗深蓝线</b>为多年逐月平均，反映该指标的季节性规律（旺季/淡季）。';
  wrap.appendChild(note);
  const grid = document.createElement('div');
  grid.className = 'grid';
  PAYLOAD.sheets[sheet].forEach(ind => {
    const card = document.createElement('div');
    card.className = 'card';
    const h3 = document.createElement('h3');
    h3.textContent = ind.name;
    const hint = document.createElement('div');
    hint.className = 'hint';
    hint.textContent = '覆盖年份：' + ind.years.join('、') + '（共 ' + ind.years.length + ' 年）';
    const chartDiv = document.createElement('div');
    chartDiv.className = 'chart';
    card.appendChild(h3); card.appendChild(hint); card.appendChild(chartDiv);
    grid.appendChild(card);
    const ec = echarts.init(chartDiv);
    const series = ind.series.map((s, i) => ({
      name: s.year + '年',
      type: 'line',
      data: s.data,
      showSymbol: false,
      lineStyle: { width: 1, color: palette[i % palette.length], opacity: 0.65 },
      emphasis: { lineStyle: { width: 2.5, opacity: 1 } },
      z: 2
    }));
    series.push({
      name: '多年均值',
      type: 'line',
      data: ind.seasonal,
      showSymbol: false,
      lineStyle: { width: 3, color: '#0f4c81' },
      z: 5,
      areaStyle: { color: 'rgba(15,76,129,0.08)' }
    });
    ec.setOption({
      tooltip: { trigger: 'axis', confine: true },
      legend: {
        type: 'scroll', top: 0,
        textStyle: { fontSize: 10 },
        data: ind.series.map(s => s.year + '年').concat(['多年均值'])
      },
      grid: { left: 8, right: 12, top: 36, bottom: 8, containLabel: true },
      xAxis: { type: 'category', data: MONTHS, boundaryGap: false, axisLabel: { fontSize: 10 } },
      yAxis: { type: 'value', scale: true, axisLabel: { fontSize: 10 } },
      series: series
    });
    window.addEventListener('resize', () => ec.resize());
  });
  wrap.appendChild(grid);
  app.appendChild(wrap);
});
</script>
</body>
</html>'''

html = html.replace('__DATA__', data_json)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(html)

total = sum(len(v) for v in sheets_data.values())
print('生成完成：', OUT)
print('各 sheet 指标数：', {k: len(v) for k, v in sheets_data.items()})
print('图表总数：', total)
