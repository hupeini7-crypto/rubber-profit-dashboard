# -*- coding: utf-8 -*-
"""渲染增强版 HTML：保留季节性图，新增核心指标历史走势、上游原料分流分析、下游轮胎生产传导分析。
依赖同目录 echarts.min.js（本地引用，离线可用）。"""
import json

with open(r'c:\Users\Lenovo\WorkBuddy\20261021112446\site_data.json', encoding='utf-8') as f:
    P = json.load(f)

DATA_JSON = json.dumps(P, ensure_ascii=False)

HTML = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>天然橡胶产业链利润分析 · 升级版</title>
<style>
  * { box-sizing: border-box; }
  body { margin:0; font-family:"Microsoft YaHei",Arial,sans-serif; background:#eef1f5; color:#1f2937; }
  header { background:linear-gradient(135deg,#0f4c81,#1f7a8c); color:#fff; padding:22px 30px; }
  header h1 { margin:0; font-size:23px; font-weight:700; }
  header p { margin:6px 0 0; opacity:.88; font-size:12.5px; }
  .wrap { margin:20px 28px; }
  .sec { background:#fff; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,.05); padding:18px 20px; margin-bottom:22px; }
  .sec h2 { font-size:18px; border-left:5px solid #1f7a8c; padding-left:11px; margin:0 0 14px; color:#0f4c81; }
  .note { font-size:12px; color:#64748b; line-height:1.7; background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:10px 12px; }
  .disclaim { font-size:11.5px; color:#94a3b8; }
  .cards { display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:14px; margin-bottom:6px; }
  .kpi { border-radius:10px; padding:14px; color:#fff; }
  .kpi.strong { background:linear-gradient(135deg,#0e7c66,#1fae8a); }
  .kpi.weak { background:linear-gradient(135deg,#a83232,#d35b5b); }
  .kpi.mid { background:linear-gradient(135deg,#2563a8,#3b82c4); }
  .kpi .lab { font-size:12px; opacity:.9; }
  .kpi .val { font-size:22px; font-weight:700; margin:4px 0; }
  .kpi .sub { font-size:11.5px; opacity:.92; }
  .mov { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:8px; }
  .movbox { background:#f8fafc; border:1px solid #e2e8f0; border-radius:9px; padding:10px 12px; }
  .movbox h4 { margin:0 0 6px; font-size:13px; }
  .movbox .up { color:#0e7c66; } .movbox .down { color:#a83232; }
  .movbox li { font-size:12px; margin:3px 0; }
  .diverge { margin-top:12px; padding:12px 14px; border-radius:9px; background:#fff7ed; border:1px solid #fed7aa; font-size:13px; line-height:1.7; }
  .drivers { margin-top:12px; }
  .drivers table { width:100%; border-collapse:collapse; font-size:12px; }
  .drivers th,.drivers td { border:1px solid #e2e8f0; padding:6px 8px; text-align:left; vertical-align:top; }
  .drivers th { background:#f1f5f9; }
  .outlook { margin-top:12px; font-size:13px; line-height:1.85; background:#f0f9ff; border:1px solid #bae6fd; border-radius:9px; padding:12px 14px; }
  .tag { display:inline-block; font-size:10.5px; padding:1px 7px; border-radius:10px; margin-left:6px; vertical-align:middle; }
  .tag.fact { background:#dcfce7; color:#166534; } .tag.guess { background:#fef9c3; color:#854d0e; }
  .tag.formula { background:#ede9fe; color:#5b21b6; } .tag.struct { background:#e0f2fe; color:#075985; }
  .tag.nr { background:#ffe4e6; color:#9f1239; }
  .divflag { font-size:12px; font-weight:600; }
  .divflag.gap { color:#b45309; background:#fef3c7; border-radius:6px; padding:1px 6px; }
  .divflag.sync { color:#475569; }
  .tag.warn { background:#fee2e2; color:#991b1b; }
  .seg { margin-bottom:10px; }
  .grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(440px,1fr)); gap:16px; }
  .card { background:#fff; border:1px solid #e7edf3; border-radius:10px; padding:12px 14px 6px; }
  .card h3 { margin:0 0 2px; font-size:14px; font-weight:600; color:#334155; }
  .badge { font-size:10.5px; padding:1px 8px; border-radius:9px; margin-left:6px; }
  .b-profit{background:#fee2e2;color:#991b1b;} .b-income{background:#dcfce7;color:#166534;}
  .b-price{background:#e0f2fe;color:#075985;} .b-spread{background:#fef3c7;color:#92400e;} .b-fx{background:#ede9fe;color:#5b21b6;}
  .b-operate{background:#fce7f3;color:#9d174d;}
  .core-star{color:#f59e0b; margin-left:4px;}
  .stat { width:100%; border-collapse:collapse; margin:8px 0 4px; font-size:12px; }
  .stat td { padding:4px 6px; border-bottom:1px solid #f1f5f9; }
  .stat td:first-child { color:#64748b; width:42%; }
  .up { color:#0e7c66; font-weight:600; } .down { color:#a83232; font-weight:600; } .neutral{color:#475569;}
  .anom { color:#a83232; font-weight:600; }
  .analysis { font-size:12px; color:#334155; line-height:1.65; margin:4px 0; background:#f8fafc; border-radius:6px; padding:7px 9px; }
  .caliber { font-size:11px; color:#94a3b8; line-height:1.6; margin-top:2px; }
  .dq { font-size:11px; color:#991b1b; background:#fee2e2; border-radius:6px; padding:5px 8px; margin-top:4px; }
  .chart { width:100%; height:300px; margin-top:4px; }
  .hist-chart { width:100%; height:230px; margin-top:6px; }
  .hist-cap { font-size:11px; color:#475569; background:#f1f5f9; border-radius:6px; padding:4px 8px; margin-top:4px; }
  .seas-cap { font-size:11px; color:#0f4c81; background:#f1f5f9; border-radius:6px; padding:5px 8px; margin-top:4px; }
  .controls { display:flex; flex-wrap:wrap; gap:12px; align-items:center; margin-bottom:12px; }
  .controls label { font-size:13px; }
  select, .btn { font-size:13px; padding:5px 9px; border-radius:7px; border:1px solid #cbd5e1; background:#fff; }
  .btn { cursor:pointer; background:#1f7a8c; color:#fff; border:none; }
  .btn.active { background:#0f4c81; }
  .lk-grid { display:grid; grid-template-columns:1.4fr 1fr; gap:16px; }
  @media(max-width:900px){ .lk-grid{grid-template-columns:1fr;} .mov{grid-template-columns:1fr;} }
  .lk-chart { width:100%; height:340px; }
  .corr-box { background:#f8fafc; border:1px solid #e2e8f0; border-radius:9px; padding:12px 14px; font-size:13px; line-height:1.8; }
  .corr-box b { color:#0f4c81; }
  .concl { font-size:12px; color:#334155; background:#fff7ed; border:1px solid #fed7aa; border-radius:8px; padding:8px 10px; margin-top:8px; line-height:1.7; }
  footer { text-align:center; color:#94a3b8; font-size:11.5px; padding:20px; line-height:1.7; }
</style>
</head>
<body>
<header>
  <h1>天然橡胶产业链利润分析 · 升级版</h1>
  <p id="hd-sub"></p>
</header>
<div class="wrap">

  <!-- 一、产业链利润观察 -->
  <div class="sec">
    <h2>产业链利润观察 <span class="tag fact">事实</span><span class="tag guess">推测</span></h2>
    <div class="cards" id="kpi-cards"></div>
      <div class="mov">
      <div class="movbox"><h4 class="up">近期改善 Top3（历史分位变化，跨指标可比）</h4><ul id="mov-up"></ul></div>
      <div class="movbox"><h4 class="down">近期恶化 Top3（历史分位变化，跨指标可比）</h4><ul id="mov-down"></ul></div>
    </div>
    <div id="diverge"></div>
    <div class="drivers">
      <h4 style="font-size:13px;margin:14px 0 6px;color:#334155;">驱动指标变化对照（描述性，非归因、非因果）</h4>
      <table id="driver-tbl"><thead><tr><th>核心利润指标</th><th>单位</th><th>历史分位</th><th>近20日变化</th><th>主要对照指标（近20日变化与方向，描述性）</th></tr></thead><tbody></tbody></table>
    </div>
    <div class="outlook" id="outlook"></div>
    <div id="upstream-spread"></div>
    <div id="downstream-op"></div>
  </div>

  <!-- 二、指标分析 + 季节性图 -->
  <div class="sec" id="ind-sec">
    <h2>指标分析（分类 · 最新值 · 分位 · 历史走势 · 季节性）</h2>
    <div class="note" id="ind-note"></div>
    <div id="segments"></div>
  </div>

  <!-- 三、盘面 vs 产业背离扫描 -->
  <div class="sec">
    <h2>盘面 vs 产业背离扫描 <span class="tag nr">盘面联动</span><span class="tag guess">推测</span></h2>
    <div class="note" id="div-note"></div>
    <div class="drivers" id="div-scan"></div>
  </div>

  <!-- 四、联动分析 -->
  <div class="sec">
    <h2>产业链联动分析（结构响应 · 盘面联动）</h2>
    <div class="note">联动基于现有数据；频率切换为日度向下聚合（周=当周有效观测均值，月=当月有效观测均值），按共同时间窗对齐，并展示匹配覆盖率、有效样本量 n、统计周期与数据截至日，绝不将低频数据伪造为日度。涉及<b>周度经营指标</b>（产能利用率/产量/库存天数）的对，默认<b>按周对齐</b>。<b>切换“对比方式（水平值/差分/涨跌幅）”或频率时，相关系数 r、有效样本量 n、覆盖率、领先滞后与自动结论均会同步重新计算（在 Python 端按所选模式与频率真实重算，非沿用）。</b>相关性/领先滞后仅基于真实匹配样本。<b>结构响应型</b>可解读产业内部传导规律（利润→开工→产量→库存等）；<b>盘面联动型</b>研究产业指标（收益/利润/价差/开工/产量/库存）与 NR 盘面价格的关系，相关性≠因果、不构成价格预测。<span class="tag struct">结构响应</span><span class="tag nr">盘面联动</span></div>
    <div class="controls">
      <label>频率：</label>
      <button class="btn active" data-f="daily" onclick="setFreq('daily')">日度</button>
      <button class="btn" data-f="weekly" onclick="setFreq('weekly')">周度</button>
      <button class="btn" data-f="monthly" onclick="setFreq('monthly')">月度</button>
      <label>对比方式：</label>
      <button class="btn active" data-m="level" onclick="setMode('level')">水平值</button>
      <button class="btn" data-m="diff" onclick="setMode('diff')">差分(变化额)</button>
      <button class="btn" data-m="return" onclick="setMode('return')">涨跌幅(%)</button>
      <label>联动对：</label>
      <select id="pair-sel" onchange="onPairChange()"></select>
    </div>
    <div class="lk-grid">
      <div><div class="lk-chart" id="lk-line"></div></div>
      <div>
        <div class="corr-box" id="lk-corr"></div>
        <div class="lk-chart" id="lk-lag" style="height:260px;"></div>
      </div>
    </div>
  </div>

  <div class="sec">
    <h2>数据口径与免责说明</h2>
    <div class="note" id="foot-note"></div>
  </div>

</div>
<footer id="footer"></footer>
<script src="echarts.min.js"></script>
<script>
const P = __DATA__;
const MONTHS = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
const CATNAME = {profit:'理论利润',income:'实际收益',price:'价格',spread:'价差',fx:'汇率',operate:'经营指标'};

document.getElementById('hd-sub').textContent =
  '日度截至 '+P.meta.asof_daily+' ｜ 周度(经营)截至 '+P.meta.asof_weekly+' ｜ 覆盖 2015–2026 ｜ 生成于 '+P.meta.generated+' ｜ 共 '+countInd()+' 个指标、'+P.linkage.length+' 组联动对 ｜ 数据源已刷新';

function countInd(){let n=0;for(const k in P.segments)n+=P.segments[k].indicators.length;return n;}

/* ---------- 摘要 ---------- */
function renderSummary(){
  const s=P.summary;
  const cards=document.getElementById('kpi-cards');
  cards.innerHTML =
    kpi('strong','当前利润最强环节', s.strongest.name, '历史分位 '+s.strongest.hist_q+'% ｜ '+s.strongest.latest+' '+s.strongest.unit)+
    kpi('weak','当前利润最弱环节', s.weakest.name, '历史分位 '+s.weakest.hist_q+'% ｜ '+s.weakest.latest+' '+s.weakest.unit)+
    kpi('mid','核心利润/收益指标数', s.core.length, '含上游收益/加工利润、中游盘面利润、下游轮胎利润（★）');
  const up=document.getElementById('mov-up'), down=document.getElementById('mov-down');
  up.innerHTML = s.movers_up.map(d=>moverLi(d,true)).join('');
  down.innerHTML = s.movers_down.map(d=>moverLi(d,false)).join('');
  if(s.divergence && s.divergence[0]){
    document.getElementById('diverge').innerHTML = '<div class="diverge"><b>上下游利润背离：</b>'+s.divergence[0]+
      ' <span class="tag guess">推测</span><br><span class="disclaim">分位差 '+s.divergence[1]+' 个百分点（上游 '+qUp()+'% vs 下游 '+qDown()+'%）；该背离已持续约 '+s.divergence[2]+' 周（自 '+s.divergence[3]+'）。</span></div>';
  }
  const tb=document.querySelector('#driver-tbl tbody');
  tb.innerHTML = s.core.map(d=>{
    const dd = s.driver_details[d.name]||{items:[]};
    const drv = dd.items.map(it=> it.dname+' '+(it.dchg20>0?'+':'')+it.dchg20+' '+it.unit+' → <b>'+it.effect+'</b>').join('；');
    const chgTxt = (d.chg20==null)?'—':((d.chg20>0?'+':'')+d.chg20+' '+d.unit+(d.show20?(' ('+(d.chg20p>0?'+':'')+d.chg20p+'%)'):' （%不适用）'));
    const chgCls = d.chg20>0?'up':(d.chg20<0?'down':'neutral');
    return '<tr><td>'+d.name+(d.unconfirmed?' <span class="tag warn">口径待确认</span>':'')+'</td><td>'+d.unit+'</td><td>'+d.hist_q+'%</td><td class="'+chgCls+'">'+chgTxt+'</td><td>'+drv+'</td></tr>';
  }).join('');
  document.getElementById('outlook').innerHTML = buildOutlook();
  renderUpstreamSpread();
  renderDownstreamOp();
  renderOctWindow();
}
function moverLi(d,improving){
  const cls = d.chg20>0?'up':'down';
  const pct = d.show20 ? (' ('+(d.chg20p>0?'+':'')+d.chg20p+'%)') : '（%不适用）';
  const qTxt = d.chg20_q==null?'—':((d.chg20_q>0?'+':'')+d.chg20_q+' 分位');
  return '<li>'+d.name+(d.unconfirmed?' <span class="tag warn">待确认</span>':'')+'：历史分位变化 <span class="'+cls+'">'+qTxt+'</span>；原始变化 <span class="'+cls+'">'+(d.chg20==null?'—':((d.chg20>0?'+':'')+d.chg20+' '+d.unit))+'</span>'+pct+'</li>';
}
function kpi(cls,lab,val,sub){return '<div class="kpi '+cls+'"><div class="lab">'+lab+'</div><div class="val">'+val+'</div><div class="sub">'+sub+'</div></div>';}
function qUp(){const d=P.summary.core.find(x=>x.name==='国内胶农收益');return d?d.hist_q:'?';}
function qDown(){const d=P.summary.core.find(x=>x.name==='全钢轮胎利润');return d?d.hist_q:'?';}
function fmt(v){return v==null?'—':v;}

function buildOutlook(){
  const s=P.summary;
  const seg=s.seg_avg_q;
  let t='<b>研究观察（动态生成，区分事实/推测）：</b><br>';
  t+='• <span class="tag fact">事实</span>当前利润分配：最强为 <b>'+s.strongest.name+'</b>（历史 '+s.strongest.hist_q+'% 分位），最弱为 <b>'+s.weakest.name+'</b>（'+s.weakest.hist_q+'% 分位）。分环节看，上游核心收益平均处于 <b>'+seg['上游']+'%</b> 分位、中游 <b>'+seg['中游']+'%</b>、下游 <b>'+seg['下游']+'%</b> 分位（各环节以各自历史分位比较，不直接比较不同单位绝对值）。<br>';
  if(s.divergence && s.divergence[0]){
    t+='• <span class="tag guess">推测</span>上下游利润背离：'+s.divergence[0]+' 该背离已持续约 <b>'+s.divergence[2]+' 周</b>（自 '+s.divergence[3]+'），分位差 '+s.divergence[1]+' 个百分点。<br>';
  } else {
    t+='• <span class="tag fact">事实</span>当前上下游利润分配未出现显著结构性背离（分位差未超阈值）。<br>';
  }
  const mu=s.movers_up.map(d=>d.name+'（'+(d.chg20>0?'+':'')+d.chg20+'）').join('、');
  const md=s.movers_down.map(d=>d.name+'（'+(d.chg20>0?'+':'')+d.chg20+'）').join('、');
  t+='• <span class="tag fact">事实</span>近20日绝对变化：改善居前为 '+mu+'；恶化居前为 '+md+'（按绝对变化额方向判断，利润型指标 up=改善）。<br>';
  const wd=s.driver_details[s.weakest.name];
  if(wd){
    const drv=wd.items.map(it=>it.dname+' '+(it.dchg20>0?'+':'')+it.dchg20+' '+it.unit+'（'+it.effect+'）').join('；');
    t+='• <span class="tag guess">推测</span>以最弱环节 '+s.weakest.name+' 为例，其近20日变化 '+(wd.pchg20>0?'+':'')+wd.pchg20+' '+wd.unit+'，主要受：'+drv+'。<br>';
  }
  t+='<br><b>需验证的产业假设（推测，非价格方向预测）：</b><br>';
  t+='• 上游高收益通常对应割胶积极性提升，后续需验证产区物候、降雨与杯胶/胶水价格是否回落；若回落则供应释放、利润向中游让渡。<br>';
  t+='• 下游（轮胎）利润若持续低位，将压制轮胎开工与补库意愿，建议用开工率数据验证（见下方“下游轮胎生产传导分析”）。<br>';
  t+='• 中游进口价差与海外报价、汇率高度相关，美元兑人民币与近港美金报价是进口利润的关键验证指标。<br>';
  t+='<span class="disclaim">注：以上为基于现有利润与成本数据的结构性推断（结构响应型联动可作产业传导参考、不构成独立证据；盘面联动仅展示相关性≠因果）；未据利润变化直接预测 RU/NR 涨跌；实际供需仍需结合库存、开工、进出口与宏观验证。</span>';
  return t;
}

/* 上游原料分流与加工选择分析 */
function renderUpstreamSpread(){
  const sp=P.summary.upstream_spread||{};
  const el=document.getElementById('upstream-spread');
  const keys=Object.keys(sp);
  if(!keys.length){el.innerHTML='';return;}
  let h='<h4 style="font-size:13.5px;margin:16px 0 8px;color:#0f4c81;">上游原料分流与加工选择分析 <span class="tag fact">事实</span><span class="tag guess">推测</span></h4>';
  h+='<div class="note" style="margin-bottom:8px;">逻辑：胶价 → 胶农收益与水杯价差 → 原料分流和加工选择（浓乳/全乳 vs 标胶）。价差走阔（胶水相对杯胶更贵）提示胶农倾向交胶水制浓乳/全乳、杯胶制标胶经济性下降；收窄则相反。<b>注意统一干胶含量与价格口径</b>（浓乳按干胶折算后比较）。</div>';
  h+='<table class="drivers" style="margin-top:6px;"><thead><tr><th>价差指标</th><th>单位</th><th>最新</th><th>历史分位</th><th>近20日变化</th><th>样本/截至</th><th>解读（推测）</th></tr></thead><tbody>';
  for(const k of ['泰国水杯价差','国内制浓乳-全乳价差']){
    const v=sp[k]; if(!v) continue;
    const chgTxt=(v.chg20==null)?'—':((v.chg20>0?'+':'')+v.chg20+' '+v.unit+(v.show20?(' ('+(v.chg20p>0?'+':'')+v.chg20p+'%)'):'（%不适用）'));
    let interp='';
    if(k==='泰国水杯价差'){
      interp = v.hist_q>=70?'当前处于历史高位（胶水相对杯胶偏贵），原料更倾向流向浓乳/全乳加工，标胶原料杯胶相对收紧。':(v.hist_q<=30?'处于历史低位（胶水相对杯胶便宜），杯胶制标胶经济性相对占优。':'处于历史中位，原料分流信号中性。');
    } else {
      interp = v.latest<0?('浓乳相对全乳深度贴水（负值，历史分位'+v.hist_q+'%），浓乳加工经济性占优、全乳承压；需核对口径。'):('浓乳相对全乳升水（历史分位'+v.hist_q+'%）。');
    }
    h+='<tr><td>'+k+'</td><td>'+v.unit+'</td><td>'+v.latest+'</td><td>'+v.hist_q+'%</td><td>'+chgTxt+'</td><td>'+v.n+' / '+v.asof+'</td><td>'+interp+'</td></tr>';
  }
  h+='</tbody></table>';
  el.innerHTML=h;
}

/* 下游轮胎生产传导分析 */
function renderDownstreamOp(){
  const op=P.summary.downstream_op||{};
  const keys=Object.keys(op);
  const el=document.getElementById('downstream-op');
  if(!keys.length){el.innerHTML='';return;}
  let h='<h4 style="font-size:13.5px;margin:16px 0 8px;color:#0f4c81;">下游轮胎生产传导分析 <span class="tag fact">事实</span><span class="tag guess">推测</span></h4>';
  h+='<div class="note" style="margin-bottom:8px;">逻辑：原料成本 → 轮胎利润 → 产能利用率与产量 → 库存可用天数。分别看全钢/半钢。判断利润改善是否伴随开工提升、产量增加、库存压力变化。<b>注意全国产能利用率与山东产量/库存的地区口径差异</b>；相关性≠因果。</div>';
  h+='<div class="mov">';
  for(const t of ['全钢','半钢']){
    const d=op[t]; if(!d) continue;
    const pr=d.profit||{}, ut=d.util||{}, ou=d.out||{}, iv=d.inv||{};
    const prChg=(pr.chg20==null)?'—':((pr.chg20>0?'+':'')+pr.chg20+' '+pr.unit+(pr.show20?(' ('+(pr.chg20p>0?'+':'')+pr.chg20p+'%)'):'（%不适用）'));
    const utChg=(ut.chg==null)?'—':((ut.chg>0?'+':'')+ut.chg);
    const ouChg=(ou.chg==null)?'—':((ou.chg>0?'+':'')+ou.chg);
    const ivChg=(iv.chg==null)?'—':((iv.chg>0?'+':'')+iv.chg);
    let trans='';
    const profitUp = pr.chg20!=null && pr.chg20>0;
    const utilUp = ut.chg!=null && ut.chg>0;
    const invUp = iv.chg!=null && iv.chg>0;
    if(profitUp && utilUp && !invUp) trans='利润改善伴随开工率上升、库存天数未升，传导较顺畅。';
    else if(profitUp && utilUp && invUp) trans='利润改善且开工/产量回升，但库存天数同步走高，提示产成品有积压风险。';
    else if(profitUp && !utilUp) trans='利润改善但开工率未同步回升，传导暂未顺畅，或受需求/检修制约。';
    else if(!profitUp && utilUp) trans='开工率回升但利润未改善，或受成本回落/让利促销驱动。';
    else trans='利润与开工变化方向需结合近期数据观察。';
    h+='<div class="movbox"><h4>'+t+'轮胎</h4>'+
       '<table class="drivers" style="margin-top:4px;"><tbody>'+
       '<tr><td>轮胎利润</td><td>'+fmt(pr.latest)+' '+pr.unit+'</td><td>历史分位 '+fmt(pr.hist_q)+'%</td><td>近20日 '+prChg+'</td></tr>'+
       '<tr><td>产能利用率(中国)</td><td>'+fmt(ut.latest)+' '+ut.unit+'</td><td>近4周 '+utChg+'</td><td>截至 '+fmt(ut.asof)+'</td></tr>'+
       '<tr><td>产量(山东)</td><td>'+fmt(ou.latest)+' '+ou.unit+'</td><td>近4周 '+ouChg+'</td><td>截至 '+fmt(ou.asof)+'</td></tr>'+
       '<tr><td>库存可用天数(山东)</td><td>'+fmt(iv.latest)+' '+iv.unit+'</td><td>近4周 '+ivChg+'</td><td>截至 '+fmt(iv.asof)+'</td></tr>'+
       '</tbody></table>'+
       '<div class="analysis" style="margin-top:4px;">传导观察（推测，非因果）：'+trans+'</div></div>';
  }
  h+='</div>';
  el.innerHTML=h;
}

/* 轮胎指标 · 国庆(10月)窗口与往年同期对比 */
function renderOctWindow(){
  const arr=P.summary.oct_window||[];
  const el=document.getElementById('downstream-op');
  if(!arr.length) return;
  const note=P.summary.oct_note||'';
  let h='<h4 style="font-size:13.5px;margin:16px 0 8px;color:#0f4c81;">轮胎指标 · 国庆(10月)窗口与往年同期对比 <span class="tag fact">事实对照</span></h4>';
  h+='<div class="note" style="margin-bottom:8px;">逻辑：检查 10月9日前后轮胎指标波动是否受节假日与统计周期影响。下表取每年 10/1–10/15 区间均值，与 2023/2024/2025 同年窗口均值比较。'+note+'</div>';
  h+='<table class="drivers" style="margin-top:6px;"><thead><tr><th>指标</th><th>单位</th><th>2023</th><th>2024</th><th>2025</th><th>2026(10/1–15)</th><th>往年均值</th><th>差值</th><th>信号</th></tr></thead><tbody>';
  for(const o of arr){
    const v=o.values;
    const cell=y=> (y in v)? v[y].toFixed(1):'—';
    const dcls = o.delta>0?'up':'down';
    h+='<tr><td>'+o.name+'</td><td>'+o.unit+'</td>'+
       '<td>'+cell(2023)+'</td><td>'+cell(2024)+'</td><td>'+cell(2025)+'</td>'+
       '<td><b>'+v[2026].toFixed(1)+'</b></td><td>'+o.prev_avg.toFixed(1)+'</td>'+
       '<td class="'+dcls+'">'+(o.delta>0?'+':'')+o.delta.toFixed(1)+'</td>'+
       '<td>'+o.flag+'</td></tr>';
  }
  h+='</tbody></table>';
  el.insertAdjacentHTML('beforeend', h);
}

/* ---------- 指标卡片 + 季节性图 + 历史走势 ---------- */
function renderIndicators(){
  document.getElementById('ind-note').innerHTML =
    '共 '+countInd()+' 个指标，按上游（收益/利润/价差）、中游（加工贸易/盘面）、下游（轮胎/原料/经营）分类；核心利润指标标★。全部指标均给出<b>完整历史走势</b>（<b>真实原始频率</b>：日度指标→日度点、周度经营指标→周度点，不聚合、不插值、不填0）与<b>年度季节性</b>（全部历史年度细线、多年均值、10–90%分位带、当年加粗线、零值线），并附当年均值与历史均值差值、样本年数与截至日期。每个指标另给出最新值、5/20日<b>绝对变化额</b>（优先）与百分比（仅当现值与前值均为正、不跨零时显示）、历史分位、同期季节性分位、历史均值±标准差与异常状态。';
  const root=document.getElementById('segments');
  for(const segk of ['上游','中游','下游']){
    const seg=P.segments[segk];
    const wrap=document.createElement('div'); wrap.className='seg';
    wrap.innerHTML='<h3 style="margin:14px 0 8px;color:#0f4c81;font-size:15px;">'+seg.title+'</h3>';
    const grid=document.createElement('div'); grid.className='grid';
    seg.indicators.forEach(ind=>{
      const card=document.createElement('div'); card.className='card';
      const star = ind.is_core?'<span class="core-star">★</span>':'';
      const bd='b-'+ind.cat;
      const st=ind.stats;
      const chg5 = st.chg5==null?'—':((st.chg5>0?'+':'')+st.chg5+' '+ind.unit+(st.show5?' ('+(st.chg5p>0?'+':'')+st.chg5p+'%)':' （%不适用）'));
      const chg20 = st.chg20==null?'—':((st.chg20>0?'+':'')+st.chg20+' '+ind.unit+(st.show20?' ('+(st.chg20p>0?'+':'')+st.chg20p+'%)':' （%不适用：跨零/负值）'));
      const cls5 = (ind.cat==='profit'||ind.cat==='income')?(st.chg5>0?'up':(st.chg5<0?'down':'neutral')):'neutral';
      const cls20 = (ind.cat==='profit'||ind.cat==='income')?(st.chg20>0?'up':(st.chg20<0?'down':'neutral')):'neutral';
      const anomCls = st.anomaly!=='正常区间'?'anom':'';
      const warn = ind.unconfirmed?'<span class="tag warn">口径待确认</span>':'';
      const dq = (ind.excluded_n>0)?('<div class="dq">⚠ 数据质量：该序列含 '+ind.excluded_n+' 个物理不可能异常值（如 2015 年 >100% 的产能利用率），已从统计与图表中显式排除，未静默沿用；未排除部分均值/分位带按 5–95% 截尾展示。源外部工作簿口径未独立核实，建议核对源数据。</div>'):'';
      card.innerHTML =
        '<h3>'+ind.name+star+'<span class="badge '+bd+'">'+CATNAME[ind.cat]+'</span>'+warn+'</h3>'+
        '<div class="stat"><table class="stat"><tr><td>最新值（'+st.latest_date+'）</td><td><b>'+st.latest+'</b> '+ind.unit+'</td></tr>'+
        '<tr><td>5日变化（绝对额优先）</td><td class="'+cls5+'">'+chg5+'</td></tr>'+
        '<tr><td>20日变化（绝对额优先）</td><td class="'+cls20+'">'+chg20+'</td></tr>'+
        '<tr><td>历史分位</td><td>'+st.hist_q+'%</td></tr>'+
        '<tr><td>同期季节性分位</td><td>'+(st.seas_q==null?'样本不足':st.seas_q+'%')+'</td></tr>'+
        '<tr><td>历史均值±标准差</td><td>'+st.mean+' ± '+st.sd+'</td></tr>'+
        '<tr><td>异常状态</td><td class="'+anomCls+'">'+st.anomaly+'</td></tr>'+
        '<tr><td>样本量 / 起始</td><td>'+st.n+' 个 ／ '+st.start+'</td></tr></table></div>'+
        '<div class="analysis">'+ind.analysis+'</div>'+
        '<div class="caliber">口径：'+ind.caliber+'</div>'+
        (dq?dq:'')+
        '<div class="chart"></div>';
      grid.appendChild(card);
      const chd=card.querySelector('.chart');
      setTimeout(()=>drawSeasonal(chd, ind), 0);
      // 全部指标均附加完整历史走势图（真实日度/周度原始频率）
      const hdiv=document.createElement('div'); hdiv.className='hist-chart'; card.appendChild(hdiv);
      const hcap=document.createElement('div'); hcap.className='hist-cap'; card.appendChild(hcap);
      setTimeout(()=>drawHistory(hdiv, ind, hcap), 0);
    });
    wrap.appendChild(grid); root.appendChild(wrap);
  }
}

function drawHistory(el, ind, capEl){
  const ec=echarts.init(el);
  const h=ind.history;
  const freqTxt = h.freq==='周度'?'周度':'日度';
  ec.setOption({
    tooltip:{trigger:'axis', confine:true, valueFormatter:v=>v==null?'-':v},
    grid:{left:6,right:10,top:24,bottom:20,containLabel:true},
    title:{text:'完整历史走势（真实'+freqTxt+'原始频率）', left:'center', textStyle:{fontSize:11,color:'#475569'}},
    xAxis:{type:'category', data:h.dates, axisLabel:{fontSize:9, formatter:v=>v? (freqTxt==='周度'? v.slice(0,10):v.slice(0,7)) : v}},
    yAxis:{type:'value', scale:true, axisLabel:{fontSize:9}},
    series:[{type:'line', data:h.vals, showSymbol:false, lineStyle:{width:1.5,color:'#1f7a8c'},
      areaStyle:{color:'rgba(31,122,140,0.08)'}}]
  });
  window.addEventListener('resize',()=>ec.resize());
  capEl.textContent='真实'+freqTxt+'原始序列（'+h.n+' 个观测点）｜ 缺失留空，不聚合/不插值/不填0 ｜ 截至 '+(freqTxt==='周度'?P.meta.asof_weekly:P.meta.asof_daily);
}

function drawSeasonal(el, ind){
  const ec=echarts.init(el);
  const s=ind.seasonal;
  const series=[];
  s.years.forEach(y=>{
    series.push({name:y+'年', type:'line', data:s.years_series[y], showSymbol:false,
      lineStyle:{width:1,opacity:0.28,color:'#9aa7b1'}, emphasis:{disabled:true}, z:1, silent:true});
  });
  series.push({name:'P10', type:'line', data:s.band10, stack:'band', symbol:'none', lineStyle:{opacity:0}, areaStyle:{opacity:0}, z:2, silent:true, tooltip:{show:false}});
  series.push({name:'P10–P90区间', type:'line', data:s.band90.map((v,i)=> v==null?null:(v-(s.band10[i]||v))), stack:'band', symbol:'none', lineStyle:{opacity:0}, areaStyle:{color:'rgba(31,122,140,0.12)'}, z:2, silent:true});
  series.push({name:'多年均值', type:'line', data:s.mean, showSymbol:false, lineStyle:{width:2.5,color:'#0f4c81'}, z:4});
  const anomalyPoints = s.anomalies.map(a=>({coord:[a[0]-1, s.cur_series[a[0]-1]], value:a[1], itemStyle:{color:'#e11d48'}}));
  series.push({name:s.cur_year+'年(当年)', type:'line', data:s.cur_series, showSymbol:true, symbolSize:5,
    lineStyle:{width:3,color:'#f59e0b'}, itemStyle:{color:'#f59e0b'}, z:6,
    markPoint: anomalyPoints.length?{symbol:'pin',symbolSize:34,data:anomalyPoints,label:{fontSize:9,color:'#fff'}}:undefined});
  const opt={
    tooltip:{trigger:'axis', confine:true, valueFormatter:v=>v==null?'-':v},
    legend:{type:'scroll', top:0, textStyle:{fontSize:10}, data:[s.cur_year+'年(当年)','多年均值','P10–P90区间'].concat(s.years.map(y=>y+'年'))},
    grid:{left:6,right:10,top:34,bottom:6,containLabel:true},
    xAxis:{type:'category', data:MONTHS, boundaryGap:false, axisLabel:{fontSize:10}},
    yAxis:{type:'value', scale:true, axisLabel:{fontSize:10}},
    series:series
  };
  if(s.zero){ opt.series[opt.series.length-1].markLine={silent:true,symbol:'none',lineStyle:{color:'#94a3b8',type:'dashed'},data:[{yAxis:0,label:{formatter:'零值',fontSize:10}}]}; }
  ec.setOption(opt);
  window.addEventListener('resize',()=>ec.resize());
  const histAvg = ind.stats.mean;
  const curAvg = s.cur_avg;
  const diff = (curAvg!=null && histAvg!=null)? (curAvg-histAvg):null;
  const cap = '当年均值 '+(curAvg==null?'-':curAvg)+' ｜ 历史均值 '+histAvg+' ｜ 差 '+(diff==null?'-':((diff>0?'+':'')+diff.toFixed(2)))+
              ' ｜ 历史样本 '+s.years.length+' 年 ｜ 截至 '+P.meta.asof;
  const capEl=document.createElement('div'); capEl.className='seas-cap'; capEl.textContent=cap;
  el.insertAdjacentElement('afterend', capEl);
}

/* ---------- 联动 ---------- */
let curFreq='daily', curMode='level';
function setFreq(f){
  curFreq=f;
  document.querySelectorAll('.controls .btn[data-f]').forEach(b=>b.classList.toggle('active', b.dataset.f===f));
  drawLink();
}
function setMode(m){
  curMode=m;
  document.querySelectorAll('.controls .btn[data-m]').forEach(b=>b.classList.toggle('active', b.dataset.m===m));
  drawLink();
}
function onPairChange(){
  const i=+document.getElementById('pair-sel').value;
  const l=P.linkage[i];
  const m = l.cross_zero ? 'diff' : 'level';
  curMode=m;
  document.querySelectorAll('.controls .btn[data-m]').forEach(b=>b.classList.toggle('active', b.dataset.m===m));
  const f = l.def_freq || 'daily';
  curFreq=f;
  document.querySelectorAll('.controls .btn[data-f]').forEach(b=>b.classList.toggle('active', b.dataset.f===f));
  drawLink();
}
function initPairSel(){
  const sel=document.getElementById('pair-sel');
  const struct=[], nr=[];
  P.linkage.forEach((l,i)=>(l.group==='nr'?nr:struct).push('<option value="'+i+'">'+l.id+'</option>'));
  sel.innerHTML='<optgroup label="结构响应产业联动（'+struct.length+'组）">'+struct.join('')+'</optgroup>'+
                '<optgroup label="NR盘面联动（'+nr.length+'组）">'+nr.join('')+'</optgroup>';
}

/* 盘面 vs 产业背离扫描 */
function renderDivScan(){
  const s=P.summary;
  const md=s.market_divergence||[];
  const ns=s.nr_scan;
  const note=document.getElementById('div-note');
  const el=document.getElementById('div-scan');
  if(!md.length){ note.innerHTML='当前无足够 NR 盘面与产业指标匹配样本，背离扫描暂不适用。'; el.innerHTML=''; return; }
  let t='<b>方法：</b>对 <b>'+md.length+' 组</b>盘面联动指标，比较产业指标<b>历史分位</b>与 <b>NR收盘价历史分位</b>的差（gap）；|gap|≥30 判为背离（<span class="divflag gap">产业强/盘面弱</span> 或 <span class="divflag gap">盘面强/产业弱</span>），否则视为<span class="divflag sync">同步</span>；并叠加<b>近20日方向一致性</b>判断（同向/背离/单边平稳/双双平稳）。';
  if(ns){
    t+='<br><b>NR 盘面基准：</b>NR收盘价历史分位 <b>'+ns.q+'%</b>，近20日变化 '+(ns.chg20==null?'—':((ns.chg20>0?'+':'')+ns.chg20))+(ns.chg20p!=null?'（'+(ns.chg20p>0?'+':'')+ns.chg20p+'%）':'（%不适用）')+'。';
  }
  t+='<br><span class="disclaim">背离仅提示“产业强弱与盘面定价出现分化”，不构成价格预测依据（相关性≠因果；实际验证仍需结合库存、开工、进出口与宏观）。</span>';
  note.innerHTML=t;
  let h='<table><thead><tr><th>产业指标</th><th>单位</th><th>产业历史分位</th><th>NR历史分位</th><th>分位差 gap</th><th>标志</th><th>近20日方向</th><th>自动解读（推测）</th></tr></thead><tbody>';
  md.forEach(d=>{
    const gapCls = Math.abs(d.gap)>=30 ? 'divflag gap' : 'divflag sync';
    const chg = d.dchg_ind==null?'—':((d.dchg_ind>0?'+':'')+d.dchg_ind+(d.show20 && d.chg20p_ind!=null ? '（'+(d.chg20p_ind>0?'+':'')+d.chg20p_ind+'%）' : '（%不适用）'));
    const nchg = d.dchg_nr==null?'—':((d.dchg_nr>0?'+':'')+d.dchg_nr);
    h+='<tr><td>'+d.name+'</td><td>'+d.unit+'</td><td>'+d.q_ind+'%</td><td>'+d.q_nr+'%</td>'+
       '<td class="'+gapCls+'">'+(d.gap>0?'+':'')+d.gap+'</td><td class="'+gapCls+'">'+d.flag+'</td>'+
       '<td>'+(d.dir||'—')+'<br><span style="font-size:11px;color:#64748b;">产业 '+(chg)+'<br>NR '+(nchg)+'</span></td>'+
       '<td style="font-size:12px;line-height:1.55;">'+d.text+'</td></tr>';
  });
  h+='</tbody></table>';
  el.innerHTML=h;
}
let lkLineEc, lkLagEc;
function drawLink(){
  const i=+document.getElementById('pair-sel').value;
  const l=P.linkage[i];
  const modeObj = l.modes[curMode] || l.modes['level'];
  const f = modeObj.freqs[curFreq];
  const ll = modeObj.leadlag;
  const concl = modeObj.concl;
  const corrBox=document.getElementById('lk-corr');
  const typeTag = l.group==='nr'?'<span class="tag nr">盘面联动</span>':'<span class="tag struct">结构响应</span>';
  if(!f){
    corrBox.innerHTML='<b>'+l.id+'</b> '+typeTag+'<br><b>样本不足</b><br>该频率下无足够匹配样本（n&lt;3），无法计算相关性。';
    if(lkLineEc)lkLineEc.clear(); if(lkLagEc)lkLagEc.clear(); return;
  }
  const modeTxt = curMode==='level'?'水平值':(curMode==='diff'?'差分(变化额)':'涨跌幅(%)');
  const xName=l.xlabel+' ('+l.unit_x+(curMode==='diff'?'·变化额':(curMode==='return'?'·%':''))+')';
  const yName=l.ylabel+' ('+l.unit_y+(curMode==='diff'?'·变化额':(curMode==='return'?'·%':''))+')';
  let html='<b>'+l.id+'</b> '+typeTag+'<br>X：'+l.xlabel+'（'+l.unit_x+'）<br>Y：'+l.ylabel+'（'+l.unit_y+'）<br>相关系数 r = <b>'+(f.corr==null?'—（不适用）':f.corr)+'</b> ｜ 有效样本 n='+f.n+' ｜ 匹配覆盖率 '+f.coverage+'%<br>统计周期：'+f.start+' ~ '+f.end+' ｜ 日度截至 '+P.meta.asof_daily+' ｜ 周度(经营)截至 '+P.meta.asof_weekly+' ｜ 对齐频率：'+(curFreq==='weekly'?'周度(默认)':curFreq)+'<br>当前对比（已重算）：'+modeTxt+'<br><span class="disclaim">'+l.link_caveat+'</span>';
  if(concl){ html+='<div class="concl"><b>自动分析结论（'+modeTxt+'，已重算）</b><br>'+concl+'</div>'; }
  corrBox.innerHTML=html;
  if(!lkLineEc) lkLineEc=echarts.init(document.getElementById('lk-line'));
  lkLineEc.setOption({
    tooltip:{trigger:'axis', confine:true},
    legend:{top:0, textStyle:{fontSize:11}, data:[l.xlabel,l.ylabel]},
    grid:{left:8,right:8,top:30,bottom:24,containLabel:true},
    xAxis:{type:'category', data:f.dates, axisLabel:{fontSize:9, formatter:v=>v.slice(0,7)}},
    yAxis:[{type:'value', scale:true, name:xName, nameTextStyle:{fontSize:10}, axisLabel:{fontSize:9}},
           {type:'value', scale:true, name:yName, nameTextStyle:{fontSize:10}, axisLabel:{fontSize:9}}],
    series:[
      {name:l.xlabel, type:'line', data:f.x, showSymbol:false, lineStyle:{width:1.5,color:'#0f4c81'}, yAxisIndex:0},
      {name:l.ylabel, type:'line', data:f.y, showSymbol:false, lineStyle:{width:1.5,color:'#e11d48'}, yAxisIndex:1}
    ]
  }, true);
  if(ll){
    if(!lkLagEc) lkLagEc=echarts.init(document.getElementById('lk-lag'));
    lkLagEc.setOption({
      tooltip:{trigger:'axis', confine:true, formatter:function(ps){const p=ps[0]; const n=ll.ns[p.dataIndex]; return '滞后 '+ll.lags[p.dataIndex]+' 月<br>r='+(p.value==null?'-':p.value)+'<br>样本 n='+n;}},
      grid:{left:8,right:8,top:28,bottom:24,containLabel:true},
      title:{text:'领先滞后（月，'+modeTxt+'，已重算）', left:'center', textStyle:{fontSize:11}},
      xAxis:{type:'category', data:ll.lags.map(x=>(x>0?'X领先'+x:(x<0?'Y领先'+(-x):'同步'))), axisLabel:{fontSize:9, interval:0, rotate:30}},
      yAxis:{type:'value', scale:true, axisLabel:{fontSize:9}, name:'相关系数'},
      series:[{type:'bar', data:ll.corrs.map((c,idx)=>({value:c, n:ll.ns[idx]})), itemStyle:{color:'#3b82c4'},
        markLine: (ll.best_corr==null? undefined : {silent:true, symbol:'none', data:[{yAxis:ll.best_corr, label:{fontSize:10, formatter:'峰值'+ll.best_corr+'@'+(ll.best_lag>0?'X领先'+ll.best_lag:(ll.best_lag<0?'Y领先'+(-ll.best_lag):'同步'))+'月'}}]})}]
    }, true);
  }
}

/* ---------- 页脚 ---------- */
document.getElementById('foot-note').innerHTML =
  '<b>数据来源与刷新：</b>上下游利润.xlsx（东方嘉信）。本次以<b>刷新后</b>工作簿为数据源（外部链接已用 Excel/COM 重新计算并保存副本；两个外部源文件「橡胶数据库 .xlsx」「橡胶-钢联数据源.xlsx」均存在并成功刷新；其中半钢产能利用率公式引用的「国内需求(周度数据)」在原始路径缺失，实际链接指向存在的「橡胶数据库 .xlsx」，刷新成功但源中 2015 年部分月份取数异常，已显式排除，详见数据质量）。'+
  '｜ <b>截至日期（分频）：</b>日度/价格/利润类截至 <b>'+P.meta.asof_daily+'</b>；周度经营指标（产能利用率/产量/库存）截至 <b>'+P.meta.asof_weekly+'</b>，二者频率不同、不可混用单一日期掩盖。'+
  '<br><b>分类：</b>★核心（共10个）= 国内胶农收益、泰国大胶园主日收益、泰国割胶农日收入、泰国加工厂国内生产利润、中国云南全乳交割利润、泰国水杯价差、泰混通关-国内市场价、泰标交NR盘面利润、全钢轮胎利润、半钢轮胎利润；其余为原料价格、汇率、国内制浓乳-全乳价差与下游经营指标（产能利用率/产量/库存），作辅助/联动分析。'+
  '<br><b>口径说明：</b>泰国胶农收益/割胶工收入按表列“（元/天）”疑已折算人民币；人民币兑泰铢=泰铢/人民币（上升=人民币升值）；美元兑人民币=人民币/美元（上升=美元升值）。全钢/半钢轮胎利润统一单位为<b>元/吨</b>（基于单条轮胎价格、重量、材料成本系数及固定成本假设的<b>理论利润</b>，模型口径待确认，不作实际经营利润引用）；上游两价差仅作<b>事实展示</b>，方向/干胶折算/加工成本未独立核实，不直接推导加工利润优势。'+
  '<br><b>地区口径差异：</b>产能利用率为“中国（全国）”口径，产量与库存可用天数为“山东”口径，二者不可直接等价比较。'+
  '<br><b>数据质量（异常排除）：</b>半钢轮胎产能利用率序列含 <b>7 个物理不可能异常值</b>（2015 年 1–8 月部分月份 >100%，最大 298968%），已按“产能利用率取值范围 0–100%”规则<b>从统计与图表中显式排除并标注</b>（非仅靠截尾掩盖），未排除部分均值/分位带按 5–95% 截尾。源外部工作簿口径未独立核实，建议核对源数据。'+
  '<br><b>严守规则：</b>缺失不填0、不插值；低频不伪造日度；切换对比方式/频率时相关性与领先滞后均真实重算；相关性/领先滞后仅基于真实匹配样本并标注样本量与周期；领先滞后≠因果；结论区分事实/推测/盘面联动/结构响应/事实对照；盘面联动与背离扫描不构成价格预测依据，未据利润变化直接预测RU/NR涨跌。未新增任何无数据支持的指标。';

document.getElementById('footer').innerHTML='天然橡胶产业链利润分析 · 升级版 ｜ 数据截至 '+P.meta.asof+' ｜ 由 WorkBuddy 自动生成 ｜ 共 '+countInd()+' 指标 / '+P.linkage.length+' 联动对';

/* init */
renderSummary();
renderIndicators();
renderDivScan();
initPairSel();
onPairChange();
</script>
</body>
</html>'''

HTML = HTML.replace('__DATA__', DATA_JSON)
with open(r'c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_年度季节性图表.html','w',encoding='utf-8') as f:
    f.write(HTML)
print('HTML written, size KB:', round(len(HTML)/1024,1))
