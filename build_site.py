# -*- coding: utf-8 -*-
"""
天然橡胶产业链利润分析 —— 数据处理层（第五轮重构，2026-11-03）
重构要点：
- 完整历史走势：hist_full() 按真实原始频率输出（日度指标→日度点；周度经营指标→周度点，相邻中位间隔≥5天判定），不聚合、不插值、不填0。
- 联动重构：删除 11 组公式驱动型联动（利润/价差由共有价格直接推算，高相关属结构性必然，不具独立研究价值）；
  保留 9 组结构响应产业联动（产业内部传导）；新增 12 组 NR 盘面联动（产业指标 vs NR收盘价），研究产业基本面与盘面定价关系。
- 盘面背离扫描：产业指标历史分位 vs NR收盘价历史分位（|gap|≥30 判背离），叠加近20日方向一致性判断，产出 market_divergence。
- 沿用第四轮修复：COM 刷新外部链接并保存副本（失败则报告并回退）；半钢产能利用率 >100% 异常显式排除；
  价差仅事实展示；联动三种模式×三频率 Python 端重算；跨指标排名用历史分位变化；日度/周度截至日期分显。
严守：缺失不填0/不插值；相关性≠因果；地区口径差异标注；不据产业指标预测盘面涨跌。
"""
import pandas as pd, numpy as np, json, math, re, os, time, warnings
from datetime import date as _date
warnings.filterwarnings('ignore')
np.seterr(all='ignore')

SRC_ORIG = r'D:\AAA东方嘉信\备份数据\上下游利润.xlsx'
REFRESHED = r'c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_refreshed.xlsx'

# ---------------- 外部链接刷新（COM）----------------
def refresh_workbook():
    st = dict(attempted=True, success=False, links=[], source_used=SRC_ORIG, note='')
    try:
        import win32com.client, urllib.parse
        xl = win32com.client.DispatchEx('Excel.Application')
        xl.Visible = False; xl.DisplayAlerts = False; xl.AskToUpdateLinks = False; xl.EnableEvents = False
        wb = xl.Workbooks.Open(SRC_ORIG, UpdateLinks=0)
        try:
            links = wb.LinkSources(1) or []
        except Exception:
            links = []
        for lnk in links:
            p = lnk
            if isinstance(p, str) and p.lower().startswith('file:'):
                p = urllib.parse.unquote(p[7:])
            st['links'].append(dict(target=str(p), exists=bool(os.path.exists(p))))
        try:
            wb.RefreshAll(); time.sleep(8); xl.CalculateFullRebuild(); st['success'] = True
        except Exception as e:
            st['note'] += ' RefreshAll异常:%r;' % e
        try:
            wb.SaveAs(REFRESHED, FileFormat=51); st['source_used'] = REFRESHED
        except Exception as e:
            st['note'] += ' SaveAs异常:%r;' % e
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        xl.Quit()
    except Exception as e:
        st['note'] += ' COM不可用:%r;' % e
        if os.path.exists(REFRESHED):
            st['source_used'] = REFRESHED; st['note'] += ' 使用已存在的刷新副本;'
        else:
            st['source_used'] = SRC_ORIG; st['note'] += ' 使用原始文件;'
    # 决定是否使用刷新副本
    if st['source_used'] == REFRESHED and not os.path.exists(REFRESHED):
        st['source_used'] = SRC_ORIG
    return st

REFRESH_STATUS = refresh_workbook()
SRC = REFRESH_STATUS['source_used']

# ---------------- 指标元数据 ----------------
META = {
 ('上游','国内胶农收益'): dict(cat='income', unit='元/天', is_core=True, better='up', unconfirmed=False,
      caliber='国内胶农割胶售胶的日度净收益（实际收益；单位与公式依字段名与量纲推断，具体以数据源为准）'),
 ('上游','泰国大胶园主日收益'): dict(cat='income', unit='元/天', is_core=True, better='up', unconfirmed=False,
      caliber='泰国大型胶园胶农日度收益（原表列名已含“（元/天）”，疑已折算为人民币；若实际为泰铢请校正）'),
 ('上游','泰国割胶农日收入'): dict(cat='income', unit='元/天', is_core=True, better='up', unconfirmed=False,
      caliber='泰国割胶工人日度劳务收入（原表列名已含“（元/天）”，疑已折算为人民币；若实际为泰铢请校正）'),
 ('上游','泰国加工厂国内生产利润'): dict(cat='profit', unit='元/吨', is_core=True, better='up', unconfirmed=False,
      caliber='泰国橡胶加工厂在国内销售环节的理论生产利润（理论利润，推断）'),
 ('上游','中国云南全乳交割利润'): dict(cat='profit', unit='元/吨', is_core=True, better='up', unconfirmed=False,
      caliber='云南全乳胶交割环节理论利润（理论利润，推断）'),
 ('上游','泰国杯胶价格'): dict(cat='price', unit='泰铢/公斤(推断)', is_core=False, better='neutral', unconfirmed=False,
      caliber='泰国杯胶（原料）价格，上游收益的成本/售价端辅助指标；单位按泰国市场惯例推断为泰铢/公斤'),
 ('上游','人民币兑泰铢'): dict(cat='fx', unit='泰铢/人民币', is_core=False, better='neutral', unconfirmed=False,
      caliber='汇率：数值=1元人民币可兑换的泰铢数。上升=人民币升值/泰铢走弱（泰国收益折算人民币增加）；下降=相反'),
 ('上游','国内胶水价格'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False,
      caliber='国内胶水（原料）价格，上游收益与交割利润的成本端辅助指标'),
 ('上游','泰国水杯价差'): dict(cat='spread', unit='泰铢/公斤', is_core=True, better='neutral', unconfirmed=False,
      caliber='事实展示：泰国胶园“水”（胶水）与“杯”（杯胶）价差=胶水价-杯胶价（泰铢/公斤）。方向、干胶含量折算与加工成本未经独立核实，仅作事实呈现，不直接推导加工利润优势。'),
 ('上游','国内制浓乳-全乳价差'): dict(cat='spread', unit='元/吨', is_core=False, better='neutral', unconfirmed=False,
      caliber='事实展示：国内浓乳与全乳价差（元/吨），可为负。比较时需将浓乳按干胶含量折算后对齐口径；方向/干胶折算/加工成本未独立核实，仅作事实呈现，不直接推导加工利润优势。'),

 ('中游','泰混通关-国内市场价'): dict(cat='spread', unit='元/吨', is_core=True, better='neutral', unconfirmed=False,
      caliber='事实展示：泰混通关后国内市场报价与通关价的价差代理（元/吨，历史上曾为负）。方向/折算口径未独立核实，仅作事实呈现。'),
 ('中游','泰标交NR盘面利润'): dict(cat='profit', unit='元/吨', is_core=True, better='up', unconfirmed=False,
      caliber='泰标交割至NR盘面理论利润（理论利润，推断）'),
 ('中游','近港美金泰混'): dict(cat='price', unit='美元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='近港美金泰混海外报价'),
 ('中游','美元兑人民币'): dict(cat='fx', unit='人民币/美元', is_core=False, better='neutral', unconfirmed=False, caliber='汇率：数值=1美元可兑换的人民币数。上升=美元升值/人民币走弱（进口成本上升）；下降=相反'),
 ('中游','近港美金泰标'): dict(cat='price', unit='美元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='近港美金泰标海外报价'),
 ('中游','泰混现货价格'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='泰混国内现货价格'),
 ('中游','NR收盘价'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='NR（20号胶）期货收盘价'),

 ('下游','全钢轮胎利润'): dict(cat='profit', unit='元/吨', is_core=True, better='up', unconfirmed=True,
      caliber='全钢轮胎理论生产利润（元/吨）。基于单条轮胎价格、重量、材料成本系数及固定成本假设计算的理论利润；模型口径未经验证，不应作为实际经营利润直接引用。'),
 ('下游','半钢轮胎利润'): dict(cat='profit', unit='元/吨', is_core=True, better='up', unconfirmed=True,
      caliber='半钢轮胎理论生产利润（元/吨）。基于单条轮胎价格、重量、材料成本系数及固定成本假设计算的理论利润；模型口径未经验证，不应作为实际经营利润直接引用。'),
 ('下游','全钢轮胎：600花纹：12R22.5-16PR：市场价：山东：玲珑轮胎'): dict(cat='price', unit='元/条', is_core=False, better='neutral', unconfirmed=False, caliber='全钢轮胎市场价（成品价）'),
 ('下游','半钢轮胎：990花纹：205/55R16：批发价格：山东：森麒麟轮胎'): dict(cat='price', unit='元/条', is_core=False, better='neutral', unconfirmed=False, caliber='半钢轮胎批发价（成品价）'),
 ('下游','顺丁橡胶：BR9000：市场价：华东地区'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='顺丁橡胶（合成胶）价格，轮胎成本端'),
 ('下游','丁苯橡胶：1712：市场价：山东：中石化齐鲁化工'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='丁苯橡胶（合成胶）价格，轮胎成本端'),
 ('下游','帘线钢：热轧：LX82A：Φ5.5：市场价：东营：青岛特钢：青钢'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='帘线钢价格，轮胎骨架成本'),
 ('下游','锦纶帘子布：1260D：市场价：华东地区'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='锦纶帘子布价格，轮胎骨架成本'),
 ('下游','炭黑：N220：市场价：山东'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='炭黑价格，轮胎成本端'),
 ('下游','氧化锌：99.70%：汇总价格：山东'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='氧化锌价格，轮胎辅料成本'),
 ('下游','现货泰混价格'): dict(cat='price', unit='元/吨', is_core=False, better='neutral', unconfirmed=False, caliber='天然橡胶（泰混）现货价格，轮胎天胶成本端'),
 ('下游','全钢轮胎：产能利用率：中国'): dict(cat='operate', unit='%', is_core=False, better='up', unconfirmed=True,
      caliber='全钢轮胎产能利用率（中国，周度，%）。反映开工强度；全国口径，与山东产量/库存存在地区口径差异。物理范围 0–100%，超出视为异常。'),
 ('下游','全钢轮胎：库存可用天数：山东'): dict(cat='operate', unit='天', is_core=False, better='neutral', unconfirmed=True,
      caliber='全钢轮胎库存可用天数（山东，周度，天）。反映库存压力；地区口径为山东，非全国。'),
 ('下游','全钢轮胎：产量：山东'): dict(cat='operate', unit='万条', is_core=False, better='up', unconfirmed=True,
      caliber='全钢轮胎产量（山东，周度，万条）。地区口径为山东，与全国产能利用率口径不同。'),
 ('下游','半钢轮胎：产能利用率：中国'): dict(cat='operate', unit='%', is_core=False, better='up', unconfirmed=True,
      caliber='半钢轮胎产能利用率（中国，周度，%）。反映开工强度；全国口径。物理范围 0–100%；2015 年部分月份出现 >100% 的录入/取数异常（源外部工作簿口径），已从统计与图表中排除并标注。'),
 ('下游','半钢轮胎：库存可用天数：山东'): dict(cat='operate', unit='天', is_core=False, better='neutral', unconfirmed=True,
      caliber='半钢轮胎库存可用天数（山东，周度，天）。反映库存压力；地区口径为山东，非全国。'),
 ('下游','半钢轮胎：产量：山东'): dict(cat='operate', unit='万条', is_core=False, better='up', unconfirmed=True,
      caliber='半钢轮胎产量（山东，周度，万条）。地区口径为山东，与全国产能利用率口径不同。'),
}

# 联动配置（第五轮重构）：
# - 删除 11 组公式驱动型联动（利润/价差由共有价格直接推算，高相关属结构性必然，不具独立研究价值）。
# - 保留 9 组结构响应产业联动（产业内部传导）。
# - 新增 12 组 NR 盘面联动：产业指标（收益/加工利润/交割利润/进口价差/轮胎利润/开工/产量/库存）vs NR收盘价。
LINK_META = {
 # —— 结构响应产业联动（保留）——
 '跨环节·上游收益 vs 中游加工利润':('structural','上下游分属不同环节，收益与加工利润此消彼长反映利润分配，属经营/结构响应型联动，可解读产业规律。'),
 '跨环节·上游收益 vs 下游全钢利润':('structural','全链利润分配与背离，跨环节对比，属结构响应型联动，可解读产业规律。'),
 '上游·胶农收益 vs 浓乳-全乳价差':('structural','国内胶农收益与浓乳-全乳价差同受国内乳胶/干胶供需影响，属结构响应型；价差走阔（浓乳相对全乳更贵）提示原料更倾向流向浓乳/全乳加工。'),
 '下游·全钢利润 vs 产能利用率':('structural','轮胎理论利润与开工率同受终端需求与成本驱动，属结构响应型联动，可解读产业规律；注意全国产能利用率口径。'),
 '下游·全钢利润 vs 产量':('structural','利润与产量同受需求与排产影响，属结构响应型；注意山东产量与全国口径差异。'),
 '下游·全钢利润 vs 库存天数':('structural','利润改善通常利于去库，但产量投放也可能推升库存，属结构响应型；注意山东库存口径。'),
 '下游·半钢利润 vs 产能利用率':('structural','轮胎理论利润与开工率同受终端需求与成本驱动，属结构响应型联动，可解读产业规律；注意全国产能利用率口径。'),
 '下游·半钢利润 vs 产量':('structural','利润与产量同受需求与排产影响，属结构响应型；注意山东产量与全国口径差异。'),
 '下游·半钢利润 vs 库存天数':('structural','利润改善通常利于去库，但产量投放也可能推升库存，属结构响应型；注意山东库存口径。'),
 # —— NR 盘面联动（新增）：产业指标 vs NR收盘价，研究产业基本面与盘面定价关系 ——
 '盘面·国内胶农收益 vs NR收盘价':('nr','上游割胶收益与NR盘面价格：收益反映原料端景气，盘面反映远期定价，相关性提示成本端向盘面传导强度。相关性≠因果。'),
 '盘面·泰国加工利润 vs NR收盘价':('nr','泰国加工厂利润与NR盘面价格：加工利润反映加工环节盈亏，与盘面同向提示加工利润与盘面定价联动。相关性≠因果。'),
 '盘面·云南交割利润 vs NR收盘价':('nr','云南全乳交割利润与NR盘面价格：交割利润含交割品价格与成本，与NR盘面同属天胶定价体系但非同一品种，联动属结构响应。相关性≠因果。'),
 '盘面·泰混通关价差 vs NR收盘价':('nr','泰混通关价差（进口盈亏代理）与NR盘面价格：价差反映进口环节盈亏，与盘面背离提示进口利润与盘面定价分化。相关性≠因果。'),
 '盘面·全钢轮胎利润 vs NR收盘价':('nr','全钢轮胎理论利润与NR盘面价格：轮胎利润受天胶成本影响，与盘面联动反映成本→盘面→成品利润的传导链条。相关性≠因果。'),
 '盘面·半钢轮胎利润 vs NR收盘价':('nr','半钢轮胎理论利润与NR盘面价格：轮胎利润受天胶成本影响，与盘面联动反映成本→盘面→成品利润的传导链条。相关性≠因果。'),
 '盘面·全钢开工率 vs NR收盘价':('nr','全钢产能利用率（中国，周度）与NR盘面价格：开工反映实体需求景气，与盘面同向提示需求与盘面定价一致；背离提示脱节。相关性≠因果。'),
 '盘面·半钢开工率 vs NR收盘价':('nr','半钢产能利用率（中国，周度）与NR盘面价格：开工反映实体需求景气，与盘面同向提示需求与盘面定价一致；背离提示脱节。相关性≠因果。'),
 '盘面·全钢产量 vs NR收盘价':('nr','全钢轮胎产量（山东，周度）与NR盘面价格：产量反映供给投放节奏，与盘面联动提示供应与定价关系。相关性≠因果。'),
 '盘面·半钢产量 vs NR收盘价':('nr','半钢轮胎产量（山东，周度）与NR盘面价格：产量反映供给投放节奏，与盘面联动提示供应与定价关系。相关性≠因果。'),
 '盘面·全钢库存天数 vs NR收盘价':('nr','全钢轮胎库存可用天数（山东，周度）与NR盘面价格：库存反映成品压力，与盘面背离提示库存与定价脱节。相关性≠因果。'),
 '盘面·半钢库存天数 vs NR收盘价':('nr','半钢轮胎库存可用天数（山东，周度）与NR盘面价格：库存反映成品压力，与盘面背离提示库存与定价脱节。相关性≠因果。'),
}

LINK_PAIRS = [
 # —— 结构响应产业联动 ——
 ('跨环节·上游收益 vs 中游加工利润','上游','国内胶农收益','上游','泰国加工厂国内生产利润'),
 ('跨环节·上游收益 vs 下游全钢利润','上游','国内胶农收益','下游','全钢轮胎利润'),
 ('上游·胶农收益 vs 浓乳-全乳价差','上游','国内胶农收益','上游','国内制浓乳-全乳价差'),
 ('下游·全钢利润 vs 产能利用率','下游','全钢轮胎利润','下游','全钢轮胎：产能利用率：中国'),
 ('下游·全钢利润 vs 产量','下游','全钢轮胎利润','下游','全钢轮胎：产量：山东'),
 ('下游·全钢利润 vs 库存天数','下游','全钢轮胎利润','下游','全钢轮胎：库存可用天数：山东'),
 ('下游·半钢利润 vs 产能利用率','下游','半钢轮胎利润','下游','半钢轮胎：产能利用率：中国'),
 ('下游·半钢利润 vs 产量','下游','半钢轮胎利润','下游','半钢轮胎：产量：山东'),
 ('下游·半钢利润 vs 库存天数','下游','半钢轮胎利润','下游','半钢轮胎：库存可用天数：山东'),
 # —— NR 盘面联动（产业指标 vs NR收盘价）——
 ('盘面·国内胶农收益 vs NR收盘价','上游','国内胶农收益','中游','NR收盘价'),
 ('盘面·泰国加工利润 vs NR收盘价','上游','泰国加工厂国内生产利润','中游','NR收盘价'),
 ('盘面·云南交割利润 vs NR收盘价','上游','中国云南全乳交割利润','中游','NR收盘价'),
 ('盘面·泰混通关价差 vs NR收盘价','中游','泰混通关-国内市场价','中游','NR收盘价'),
 ('盘面·全钢轮胎利润 vs NR收盘价','下游','全钢轮胎利润','中游','NR收盘价'),
 ('盘面·半钢轮胎利润 vs NR收盘价','下游','半钢轮胎利润','中游','NR收盘价'),
 ('盘面·全钢开工率 vs NR收盘价','下游','全钢轮胎：产能利用率：中国','中游','NR收盘价'),
 ('盘面·半钢开工率 vs NR收盘价','下游','半钢轮胎：产能利用率：中国','中游','NR收盘价'),
 ('盘面·全钢产量 vs NR收盘价','下游','全钢轮胎：产量：山东','中游','NR收盘价'),
 ('盘面·半钢产量 vs NR收盘价','下游','半钢轮胎：产量：山东','中游','NR收盘价'),
 ('盘面·全钢库存天数 vs NR收盘价','下游','全钢轮胎：库存可用天数：山东','中游','NR收盘价'),
 ('盘面·半钢库存天数 vs NR收盘价','下游','半钢轮胎：库存可用天数：山东','中游','NR收盘价'),
]
NR_PAIRS = [p for p in LINK_PAIRS if p[0].startswith('盘面·')]
# 含周度经营指标（产能利用率/产量/库存）的对，默认按周对齐
def is_weekly_pair(xc, yc):
    return any(k in xc+yc for k in ('产能利用率','产量：','库存可用天数'))

SHEET_ORDER = ['上游','中游','下游']
SEG_TITLE = {'上游':'上游 · 原料与割胶收益','中游':'中游 · 加工、贸易与盘面','下游':'下游 · 轮胎与成品'}
OPERATE_CATS = {'operate'}

# ---------------- 读取与清洗 ----------------
xl = pd.ExcelFile(SRC)
series = {}; asof = {}
def keyname(col): return re.sub(r'（[^（）]*）', '', col).strip()
for s in SHEET_ORDER:
    df = pd.read_excel(xl, s); cols = list(df.columns); cols[0]='dt'; df.columns = cols
    df['dt'] = pd.to_datetime(df['dt'], errors='coerce'); df = df.dropna(subset=['dt']).copy()
    df['yr'] = df['dt'].dt.year; df['mon'] = df['dt'].dt.month
    for c in [x for x in df.columns if x not in ('dt','yr','mon')]:
        sub = df[['dt','yr','mon',c]].dropna(subset=[c])
        if sub.empty: continue
        ss = sub.set_index('dt')[c].sort_index()
        series[(s,keyname(c))] = ss
        asof[(s,keyname(c))] = sub['dt'].max().date()

UNIT={}; CAT={}; UNCONF={}
for (s,c),m in META.items():
    UNIT[c]=m['unit']; CAT[c]=m['cat']; UNCONF[c]=m.get('unconfirmed',False)

# ---------------- 异常排除（物理不可能值）----------------
# 产能利用率(%) 物理范围 0–100；超出者（如 2015 年半钢的 298968%）显式排除出统计与图表，并标注。
EXCL = {}
for (s,c),ss in series.items():
    if '产能利用率' in c:
        bad = ss[(ss > 100) | (ss < 0)]
        if len(bad):
            EXCL[(s,c)] = set(bad.index)
def filt(ss, key):
    e = EXCL.get(key)
    if not e: return ss
    return ss.drop(index=[d for d in e if d in ss.index])

# ---------------- 工具 ----------------
def clean(v):
    if v is None: return None
    try:
        if isinstance(v, float) and math.isnan(v): return None
    except Exception: pass
    if isinstance(v, (np.integer,)): return int(v)
    if isinstance(v, (np.floating,)): return float(v)
    if isinstance(v, (np.bool_,)): return bool(v)
    return v
def nat(x): return x is None or (isinstance(x,float) and math.isnan(x))
def robust_moments(arr):
    a = np.asarray(arr, float)
    if a.size == 0: return 0.0, 0.0, False
    med = float(np.median(a))
    lo, hi = np.percentile(a, 5), np.percentile(a, 95)
    t = a[(a >= lo) & (a <= hi)]
    return float(t.mean()), float(t.std()), False

def stats_of(ss, key):
    ss = filt(ss, key)
    vals = ss.dropna().values.astype(float)
    latest_date = ss.index.max(); latest = float(ss.loc[latest_date])
    def chg(days):
        tgt = latest_date - pd.Timedelta(days=days); prev = ss.asof(tgt)
        if nat(prev): return None, None, None
        prev = float(prev); d = latest - prev
        p = (d/prev*100) if (prev not in (0,None) and prev>0 and latest>0) else None
        return round(d,4), (round(p,2) if p is not None else None), prev
    c5,c5p,p5 = chg(5); c20,c20p,p20 = chg(20)
    hist_q = float((vals <= latest).mean()*100)
    # 历史分位变化：当前分位 vs 20日前值的分位
    chg20_q = None
    if p20 is not None:
        prev_q = float((vals <= p20).mean()*100)
        chg20_q = round(hist_q - prev_q, 1)
    mon = latest_date.month
    same_mon = ss[ss.index.month==mon].dropna().values.astype(float)
    seas_q = float((same_mon <= latest).mean()*100) if len(same_mon)>1 else None
    mean, sd, _ = robust_moments(vals)
    z = (latest-mean)/sd if sd>0 else 0
    if hist_q>=90: anomaly='偏高(历史>90%分位)'
    elif hist_q<=10: anomaly='偏低(历史<10%分位)'
    elif abs(z)>=2: anomaly='偏离均值(|Z|>=2)'
    else: anomaly='正常区间'
    return dict(
        latest=round(latest,4), latest_date=str(latest_date.date()),
        chg5=clean(c5), chg5p=clean(c5p), chg5_prev=clean(p5),
        chg20=clean(c20), chg20p=clean(c20p), chg20_prev=clean(p20), chg20_q=chg20_q,
        show5=(c5p is not None), show20=(c20p is not None),
        hist_q=round(hist_q,1), seas_q=(round(seas_q,1) if seas_q is not None else None),
        mean=round(mean,4), sd=round(sd,4), z=round(float(z),2), anomaly=anomaly,
        n=int(len(vals)), start=str(ss.index.min().date()))

def recent_change(ss, days=20):
    ss = filt(ss, None) if False else ss
    if ss is None or ss.dropna().empty: return None, None
    ss2 = ss.dropna(); ld = ss2.index.max(); latest = float(ss2.loc[ld])
    tgt = ld - pd.Timedelta(days=days); prev = ss2.asof(tgt)
    if nat(prev): return None, latest
    return round(float(latest-float(prev)),4), round(latest,4)

def seasonal_of(ss, is_zero, key):
    ss = filt(ss, key)
    tmp = ss.dropna().to_frame('v'); tmp['yr']=tmp.index.year; tmp['mon']=tmp.index.month
    pivot = tmp.pivot_table(index='mon', columns='yr', values='v', aggfunc='mean').reindex(range(1,13))
    years = [int(y) for y in pivot.columns]; cur_year = max(years)
    band10=[None if nat(v) else round(float(v),4) for v in pivot.quantile(0.10,axis=1).values]
    band90=[None if nat(v) else round(float(v),4) for v in pivot.quantile(0.90,axis=1).values]
    mean=[None if nat(v) else round(float(v),4) for v in pivot.mean(axis=1).values]
    cur_series=[None if nat(v) else round(float(v),4) for v in pivot[cur_year].values]
    b5=[None if nat(v) else round(float(v),4) for v in pivot.quantile(0.05,axis=1).values]
    b95=[None if nat(v) else round(float(v),4) for v in pivot.quantile(0.95,axis=1).values]
    anomalies=[]
    for i in range(12):
        cv=cur_series[i]
        if cv is not None and b5[i] is not None and b95[i] is not None:
            if cv>b95[i]: anomalies.append((i+1,'高于历史95%分位'))
            elif cv<b5[i]: anomalies.append((i+1,'低于历史5%分位'))
    years_series = {int(y):[None if nat(v) else round(float(v),4) for v in pivot[y].values] for y in years}
    cur_vals=[v for v in cur_series if v is not None]
    cur_avg = round(float(np.mean(cur_vals)),4) if cur_vals else None
    vv = ss.dropna().values.astype(float); rmean,_,_ = robust_moments(vv)
    return dict(years=years, cur_year=cur_year, years_series=years_series,
                band10=band10, band90=band90, mean=round(rmean,4), cur_series=cur_series,
                cur_avg=cur_avg, zero=is_zero, anomalies=anomalies)

def hist_full(ss, key):
    """完整历史走势：真实原始频率序列（日度指标→日度点；周度经营指标→周度点），不聚合、不插值。"""
    ss = filt(ss, key)
    tmp = ss.dropna()
    if tmp.empty:
        return dict(dates=[], vals=[], freq='日度', n=0)
    diffs = tmp.index.to_series().diff().dt.days
    med = float(diffs.median(skipna=True)) if len(diffs) > 1 else 1
    freq = '周度' if med >= 5 else '日度'
    dates = [str(d.date()) for d in tmp.index]
    vals = [None if nat(v) else round(float(v), 4) for v in tmp.values]
    return dict(dates=dates, vals=vals, freq=freq, n=int(len(tmp)))

# ---------------- 联动：模式 × 频率 重算 ----------------
def transform(ss, freq, mode):
    if freq=='daily': r = ss
    elif freq=='weekly': r = ss.resample('W').mean()
    else: r = ss.resample('ME').mean()
    r = r.dropna()
    if mode=='diff': r = r.diff().dropna()
    elif mode=='return': r = r.pct_change().dropna()
    return r

def corr_mode(sx, sy, freq, mode):
    a=transform(sx,freq,mode); b=transform(sy,freq,mode)
    dfj=pd.concat([a,b],axis=1); union=len(dfj); m=dfj.dropna()
    if m.shape[1]!=2 or len(m)<3: return None
    m.columns=['x','y']
    c=float(m['x'].corr(m['y']))
    cov=round(len(m)/union*100,1) if union>0 else 0
    return dict(corr=round(c,3),n=int(len(m)),coverage=cov,
                start=str(m.index.min().date()),end=str(m.index.max().date()),
                xd=[str(d.date()) for d in m.index], x=[clean(v) for v in m['x'].tolist()],
                y=[clean(v) for v in m['y'].tolist()])

def leadlag_mode(sx, sy, mode):
    a=transform(sx,'monthly',mode); b=transform(sy,'monthly',mode)
    dfj=pd.concat([a,b],axis=1).dropna(); dfj.columns=['x','y']
    if len(dfj)<10: return None
    lags=list(range(-6,7)); corrs=[]; ns=[]
    for L in lags:
        yl=dfj['y'].shift(-L); sub=pd.concat([dfj['x'],yl],axis=1).dropna(); sub.columns=['x','y']
        ns.append(int(len(sub)))
        corrs.append(round(float(sub['x'].corr(sub['y'])),3) if len(sub)>3 else None)
    best_i=max(range(len(corrs)), key=lambda i:(abs(corrs[i]) if corrs[i] is not None else -1))
    return dict(lags=lags,corrs=corrs,ns=ns,best_lag=lags[best_i],best_corr=corrs[best_i],n=int(len(dfj)))

MODE_LABEL={'level':'水平值','diff':'差分(一阶变化)','return':'涨跌幅(%)'}
def make_concl(ltype, xname, yname, mode, r, sx, sy, unitx, unity, factual):
    dchg_x, lat_x = recent_change(sx, 20); dchg_y, lat_y = recent_change(sy, 20)
    mlab = MODE_LABEL[mode]
    if ltype == 'nr':
        head='【盘面联动·%s】产业指标 %s ↔ NR盘面 %s：相关系数 r=%s。' % (mlab, xname, yname, r)
    elif ltype == 'formula':
        head='【公式/价差推导·%s】%s ↔ %s：相关系数 r=%s。' % (mlab, xname, yname, r)
    else:
        head='【结构响应·%s】%s ↔ %s：相关系数 r=%s。' % (mlab, xname, yname, r)
    fx=('%s 最新 %.2f %s，近20日变化 %s；'%(xname,lat_x,unitx,'%+.2f'%dchg_x)) if (dchg_x is not None and lat_x is not None) else ''
    fy=('%s 最新 %.2f %s，近20日变化 %s。'%(yname,lat_y,unity,'%+.2f'%dchg_y)) if (dchg_y is not None and lat_y is not None) else ''
    if ltype == 'nr':
        tail='注：相关性≠因果；产业指标与盘面价格的联动不构成价格预测依据，背离含义见上方“盘面 vs 产业背离扫描”。'
    elif factual:
        tail='事实对照（非因果、非归因）：仅呈现两者%s的统计联动；方向、干胶含量折算与加工成本未独立核实，不直接推导加工利润优势。相关性≠因果。' % mlab
    elif ('产能利用率' in yname) or ('产量' in yname) or ('库存' in yname):
        tail='描述性对照（非因果）：按“原料成本→轮胎利润→开工率与产量→库存天数”观察传导。若利润与开工率、产量同向改善且库存天数未明显上行，传导顺畅；若利润改善但库存天数走高，提示产成品积压、传导受阻。注意全国产能利用率与山东产量/库存的地区口径差异。'
    else:
        tail='注：相关性≠因果；'+('公式/价差推导型联动的高相关属结构性必然，不解读为独立产业规律。' if ltype=='formula' else '结构响应型可解读产业规律，但仍非因果。')
    return head+fx+fy+tail

linkage=[]
for pid,sa,xc,sb,yc in LINK_PAIRS:
    sx=series.get((sa,xc)); sy=series.get((sb,yc))
    if sx is None or sy is None: continue
    sx_f=filt(sx,(sa,xc)); sy_f=filt(sy,(sb,yc))
    ltype, lcaveat = LINK_META.get(pid, ('structural',''))
    factual = False
    df = 'weekly' if is_weekly_pair(xc,yc) else 'daily'
    modes={}
    for mode in ('level','diff','return'):
        freqs={}
        for f in ('daily','weekly','monthly'):
            r=corr_mode(sx_f,sy_f,f,mode)
            if r: freqs[f]=r
        ll=leadlag_mode(sx_f,sy_f,mode)
        r0=(freqs.get(df) or freqs.get('daily') or {}).get('corr','-')
        concl=make_concl(ltype,xc,yc,mode,r0,sx_f,sy_f,UNIT.get(xc,''),UNIT.get(yc,''),factual)
        modes[mode]=dict(freqs=freqs, leadlag=ll, concl=concl)
    xs=sx_f.dropna().values; ys=sy_f.dropna().values
    cross=(xs.min()<0<xs.max()) or (ys.min()<0<ys.max())
    linkage.append(dict(id=pid, group=('nr' if ltype=='nr' else 'structural'), xlabel=xc, ylabel=yc,
                        unit_x=UNIT.get(xc,''), unit_y=UNIT.get(yc,''),
                        link_type=ltype, link_caveat=lcaveat, factual=factual, cross_zero=bool(cross),
                        def_freq=df, modes=modes))

# ---------------- 盘面 vs 产业背离扫描（P1：利润/产量/开工/库存与盘面价格的背离与变化）----------------
NR_NAME='NR收盘价'
def nr_q_chg():
    sx=series.get(('中游',NR_NAME))
    if sx is None or sx.dropna().empty: return None,None,None
    st=stats_of(sx,('中游',NR_NAME))
    return st['hist_q'], st['chg20'], st['chg20p']
def div_text(name,q,nrq,gap,dchg,nrchg,dirstr,flag,unit):
    if '收益' in name or '收入' in name: tag='收益'
    elif '利润' in name: tag='利润'
    elif '价差' in name or '通关' in name: tag='价差'
    elif '产能利用率' in name: tag='开工'
    elif '产量' in name: tag='产量'
    else: tag='库存'
    base='产业指标 <b>%s</b> 历史分位 %s%% vs NR收盘价 %s%%（分位差 %+.1f）→ %s。' % (name,q,nrq,gap,flag)
    interp={
      '收益':'上游收益与盘面同向通常对应原料成本支撑强化；背离提示成本端与盘面定价分化。',
      '利润':'加工/轮胎利润与盘面背离提示环节盈亏与盘面定价分化：利润分位高而盘面弱，或为成本支撑难向下游与盘面传导；利润分位低而盘面强，盘面或已计入未来改善预期。',
      '价差':'进口/通关价差与盘面背离提示进口盈亏与盘面定价分化：价差分位高而盘面弱，进口利润兑现存疑。',
      '开工':'开工率与盘面背离提示实体景气与盘面定价脱节：开工强而盘面弱，需求好但盘面被预期/宏观压制。',
      '产量':'产量与盘面背离提示供给投放与盘面定价脱节：产量高而盘面弱，供应压力未在盘面充分定价。',
      '库存':'库存与盘面背离提示库存压力与盘面定价脱节：库存高而盘面强，盘面或定价未来去库预期。',
    }[tag]
    dir_txt=''
    if dirstr=='背离' and dchg is not None and nrchg is not None:
        dir_txt='近20日方向背离（产业指标 %s%+.2f %s，NR %s%+.0f 元/吨）。' % ('+' if dchg>0 else '', dchg, unit, '+' if nrchg>0 else '', nrchg)
    elif dirstr:
        dir_txt='近20日方向：%s。' % dirstr
    return base+interp+dir_txt+'（相关性≠因果，不构成价格预测）'
def divergence_scan():
    out=[]
    nrq,nrchg,_ = nr_q_chg()
    if nrq is None: return out
    for pid,sa,xc,sb,yc in NR_PAIRS:
        sx=series.get((sa,xc))
        if sx is None or sx.dropna().empty: continue
        st=stats_of(sx,(sa,xc))
        q=st['hist_q']; dchg=st['chg20']; show=st['show20']; dp=st['chg20p']
        gap=round(q-nrq,1)
        if gap>=30: flag='产业强/盘面弱'
        elif gap<=-30: flag='盘面强/产业弱'
        else: flag='同步'
        dirstr=None
        if dchg is not None and nrchg is not None:
            ds=1 if dchg>0 else (-1 if dchg<0 else 0); dn=1 if nrchg>0 else (-1 if nrchg<0 else 0)
            if ds!=0 and dn!=0: dirstr='同向' if ds==dn else '背离'
            elif ds==0 and dn!=0: dirstr='指标平稳/盘面%s' % ('上行' if dn>0 else '下行')
            elif ds!=0 and dn==0: dirstr='指标%s/盘面平稳' % ('上行' if ds>0 else '下行')
            else: dirstr='双双平稳'
        text=div_text(xc,q,nrq,gap,dchg,nrchg,dirstr,flag,UNIT.get(xc,''))
        out.append(dict(name=xc, unit=UNIT.get(xc,''), q_ind=round(q,1), q_nr=round(nrq,1), gap=gap,
                        dchg_ind=(round(dchg,2) if dchg is not None else None),
                        chg20p_ind=(round(dp,2) if (dp is not None and show) else None),
                        show20=bool(show),
                        dchg_nr=(round(nrchg,2) if nrchg is not None else None),
                        dir=dirstr, flag=flag, text=text))
    return out
market_divergence=divergence_scan()
# NR 盘面基准信息（供页面背离扫描区块展示）
_nrq,_nrchg,_nrchgp = nr_q_chg()
nr_scan = (dict(q=round(_nrq,1), chg20=(round(_nrchg,2) if _nrchg is not None else None),
                chg20p=(round(_nrchgp,2) if _nrchgp is not None else None)) if _nrq is not None else None)

# ---------------- 研究摘要 ----------------
core=[]
for (s,c),m in META.items():
    if m['is_core'] and (s,c) in series:
        st=stats_of(series[(s,c)],(s,c))
        core.append(dict(sheet=s,name=c,hist_q=st['hist_q'],chg20=st['chg20'],chg20p=st['chg20p'],
                         show20=st['show20'],chg20_q=st['chg20_q'],unit=m['unit'],latest=st['latest'],
                         anomaly=st['anomaly'],unconfirmed=m.get('unconfirmed',False)))
strong=max(core,key=lambda d:d['hist_q']); weak=min(core,key=lambda d:d['hist_q'])
movers=[d for d in core if d['chg20_q'] is not None]
movers_up=sorted(movers,key=lambda d:-d['chg20_q'])[:3]
movers_down=sorted(movers,key=lambda d:d['chg20_q'])[:3]

def q_of(name):
    for d in core:
        if d['name']==name: return d['hist_q']
    return None
q_up=q_of('国内胶农收益'); q_mid=q_of('泰标交NR盘面利润'); q_down=q_of('全钢轮胎利润')
def divergence_weeks(s_up,s_down,thr=30):
    a=s_up.resample('W').mean().dropna(); b=s_down.resample('W').mean().dropna()
    au=s_up.dropna().values.astype(float); bu=s_down.dropna().values.astype(float)
    idx=a.index.intersection(b.index); gap={}
    for dt in idx:
        pu=float((au<=a.loc[dt]).mean()*100); pb=float((bu<=b.loc[dt]).mean()*100); gap[dt]=pu-pb
    streak=0; since=None
    for dt in reversed(idx):
        if gap[dt]>=thr: streak+=1; since=dt
        else: break
    return streak,(str(since.date()) if since else None),thr
div_weeks=None; div_since=None
if (('上游','国内胶农收益') in series) and (('下游','全钢轮胎利润') in series):
    div_weeks,div_since,_=divergence_weeks(series[('上游','国内胶农收益')],series[('下游','全钢轮胎利润')])
if q_up is not None and q_down is not None:
    gap=q_up-q_down
    if gap>=30: divergence_text=('上游收益处于历史高位而下游轮胎利润处于低位，利润分配向上游倾斜，反映原料强、终端弱的结构性背离（推测：成本端强势挤压下游，需观察轮胎开工与替换需求验证）。')
    elif gap<=-30: divergence_text=('下游轮胎利润相对上游收益偏高，利润分配向下游倾斜（推测：需结合终端需求与轮胎提价验证）。')
    else: divergence_text=None
else: gap=None; divergence_text=None

# 驱动指标变化对照（描述性，非归因）
DRIVERS={
 '国内胶农收益':[('上游','国内胶水价格','售价端')],
 '泰国大胶园主日收益':[('上游','泰国杯胶价格','售价端'),('上游','人民币兑泰铢','折算')],
 '泰国割胶农日收入':[('上游','泰国杯胶价格','售价端'),('上游','人民币兑泰铢','折算')],
 '泰国加工厂国内生产利润':[('上游','泰国杯胶价格','成本端')],
 '中国云南全乳交割利润':[('上游','国内胶水价格','成本端'),('下游','现货泰混价格','成品/替代')],
 '泰标交NR盘面利润':[('中游','近港美金泰标','成本端'),('中游','美元兑人民币','折算'),('中游','NR收盘价','成品端')],
 '全钢轮胎利润':[('下游','现货泰混价格','天胶成本'),('下游','顺丁橡胶：BR9000：市场价：华东地区','合成胶成本'),('下游','炭黑：N220：市场价：山东','辅料成本')],
 '半钢轮胎利润':[('下游','现货泰混价格','天胶成本'),('下游','顺丁橡胶：BR9000：市场价：华东地区','合成胶成本')],
}
def role_effect(role,dc):
    if role in ('成本端','天胶成本','合成胶成本','辅料成本'): return '成本上行→模型利润承压' if dc>0 else ('成本回落→模型利润改善' if dc<0 else '持平')
    if role in ('售价端','成品端','成品/替代'): return '售价上行→模型利润改善' if dc>0 else ('售价下行→模型利润承压' if dc<0 else '持平')
    return '持平'
def fx_effect(dname,dc):
    if dname=='人民币兑泰铢': return '人民币走强→泰国收益折算增加' if dc>0 else ('人民币走弱→折算减少' if dc<0 else '持平')
    if dname=='美元兑人民币': return '人民币走弱→进口成本上升' if dc>0 else ('人民币走强→进口成本下降' if dc<0 else '持平')
    return '持平'
driver_details={}
for name,dlist in DRIVERS.items():
    key=[k for k in core if k['name']==name]
    pstat=stats_of(series[([ (s,c) for (s,c),m in META.items() if c==name][0])],([ (s,c) for (s,c),m in META.items() if c==name][0]))
    pc=pstat['chg20']; items=[]
    for sa,xc,role in dlist:
        sx=series.get((sa,xc))
        if sx is None: continue
        dc=stats_of(sx,(sa,xc))['chg20']
        if dc is None: continue
        eff=fx_effect(xc,dc) if role=='折算' else role_effect(role,dc)
        items.append(dict(dname=xc,role=role,dchg20=dc,unit=UNIT.get(xc,''),effect=eff))
    driver_details[name]=dict(pchg20=pc,unit='',items=items)
for name,dd in driver_details.items():
    dd['unit']=next((d['unit'] for d in core if d['name']==name),'')

seg_avg_q={}
for s in SHEET_ORDER:
    qs=[d['hist_q'] for d in core if d['sheet']==s]
    seg_avg_q[s]=round(float(np.mean(qs)),1) if qs else None

# 上游原料分流（价差事实展示）
def spread_of(name):
    ss=series.get(('上游',name))
    if ss is None: return None
    st=stats_of(ss,('上游',name))
    return dict(name=name,unit=UNIT.get(name,''),latest=st['latest'],hist_q=st['hist_q'],
                chg20=st['chg20'],show20=st['show20'],chg20p=st['chg20p'],anomaly=st['anomaly'],
                n=st['n'],asof=str(asof.get(('上游',name))) if ('上游',name) in asof else None)
upstream_spread={nm:spread_of(nm) for nm in ['泰国水杯价差','国内制浓乳-全乳价差'] if spread_of(nm)}

# 下游轮胎生产传导
def op_block(tire):
    out={}
    p=series.get(('下游',tire+'轮胎利润'))
    if p is not None:
        st=stats_of(p,('下游',tire+'轮胎利润'))
        out['profit']=dict(hist_q=st['hist_q'],chg20=st['chg20'],show20=st['show20'],chg20p=st['chg20p'],latest=st['latest'],unit=UNIT.get(tire+'轮胎利润',''))
    for metric,key in [('util','产能利用率：中国'),('out','产量：山东'),('inv','库存可用天数：山东')]:
        s=series.get(('下游',tire+'轮胎：'+key))
        if s is not None:
            d,lat=recent_change(s,28)
            out[metric]=dict(latest=lat,unit=UNIT.get(tire+'轮胎：'+key,''),chg=d,asof=str(s.dropna().index.max().date()))
    return out
downstream_op={t:op_block(t) for t in ['全钢','半钢']}

# 日度/周度截至日期分别
asof_daily=max([v for k,v in asof.items() if '产能利用率' not in k[1] and '产量：' not in k[1] and '库存可用天数' not in k[1]])
asof_weekly=max([v for k,v in asof.items() if ('产能利用率' in k[1]) or ('产量：' in k[1]) or ('库存可用天数' in k[1])])

# 10月9日窗口 vs 往年同期（含国庆节假日提示）
def oct_window_analysis():
    out=[]; tire=['全钢轮胎利润','半钢轮胎利润','全钢轮胎：产能利用率：中国','半钢轮胎：产能利用率：中国',
                 '全钢轮胎：产量：山东','半钢轮胎：产量：山东','全钢轮胎：库存可用天数：山东','半钢轮胎：库存可用天数：山东']
    for name in tire:
        ss=series.get(('下游',name))
        if ss is None: continue
        res={}
        for yr in [2023,2024,2025,2026]:
            sub=ss[(ss.index.year==yr)&(ss.index.month==10)&(ss.index.day>=1)&(ss.index.day<=15)]
            if len(sub)>0: res[yr]=round(float(sub.mean()),3)
        if 2026 in res and any(y<2026 for y in res):
            prev_avg=round(float(np.mean([res[y] for y in res if y<2026])),3)
            cur=res[2026]; delta=round(cur-prev_avg,3)
            out.append(dict(name=name,unit=UNIT.get(name,''),values=res,prev_avg=prev_avg,cur=cur,delta=delta,
                            flag=('偏高' if delta>0 else ('偏低' if delta<0 else '持平'))))
    note='10月1–7日为国庆假期，周度统计口径在假期所在周观测数可能减少；跨年窗口比较需注意节假日错位与统计周期差异，不能据此直接判定趋势突变。'
    return out,note
oct_window, oct_note = oct_window_analysis()

summary=dict(
    asof=str(max([v for v in asof.values()])), asof_daily=str(asof_daily), asof_weekly=str(asof_weekly),
    strongest=strong, weakest=weak, movers_up=movers_up, movers_down=movers_down,
    divergence=(divergence_text,round(gap,1) if gap is not None else None,div_weeks,div_since),
    driver_details=driver_details, seg_avg_q=seg_avg_q, core=core, upstream_spread=upstream_spread,
    downstream_op=downstream_op, oct_window=oct_window, oct_note=oct_note,
    market_divergence=market_divergence, nr_scan=nr_scan, refresh=REFRESH_STATUS)

# ---------------- 组装指标数据 ----------------
segments={}
for s in SHEET_ORDER:
    inds=[]
    for c,m in META.items():
        if c[0]!=s: continue
        if (s,c[1]) not in series: continue
        ss=series[(s,c[1])]; st=stats_of(ss,(s,c[1]))
        is_zero=m['cat'] in ('profit','income','spread')
        seas=seasonal_of(ss,is_zero,(s,c[1])); hist=hist_full(ss,(s,c[1]))
        excl_n=len(EXCL.get((s,c[1]),set()))
        lvl='高位' if st['hist_q']>=70 else ('低位' if st['hist_q']<=30 else '中位')
        trend='近期上行' if (st['chg20'] is not None and st['chg20']>0) else ('近期下行' if (st['chg20'] is not None and st['chg20']<0) else '近期平稳')
        chg20_txt='%.2f %s'%(st['chg20'],m['unit']) if st['chg20'] is not None else '样本不足'
        pct_txt=('（%.1f%%）'%st['chg20p'] if st['show20'] else '（%不适用：跨零/负值）')
        dq=''
        if excl_n: dq='【数据质量】该序列含 %d 个物理不可能异常值（如 2015 年 >100%% 的产能利用率），已从统计与图表中排除，未静默沿用。'%excl_n
        analysis='最新 %.2f %s，处于历史%s（%s分位），近20日绝对变化 %s%s，历史分位变化 %s；%s；同期季节性分位 %s。%s%s' % (
            st['latest'],m['unit'],lvl,st['hist_q'],chg20_txt,pct_txt,
            ('%+.1f'%st['chg20_q'] if st['chg20_q'] is not None else '样本不足'),
            trend,('%.0f%%'%st['seas_q'] if st['seas_q'] is not None else '样本不足'),
            ('异常提示：'+st['anomaly'] if st['anomaly']!='正常区间' else '处于正常区间'), dq)
        inds.append(dict(name=c[1],cat=m['cat'],unit=m['unit'],is_core=m['is_core'],better=m['better'],
                         unconfirmed=m.get('unconfirmed',False),caliber=m['caliber'],stats=st,
                         outlier=excl_n>0,excluded_n=excl_n,seasonal=seas,history=hist,analysis=analysis))
    segments[s]=dict(title=SEG_TITLE[s],indicators=inds)

payload=dict(meta=dict(asof=summary['asof'],asof_daily=summary['asof_daily'],asof_weekly=summary['asof_weekly'],
             generated=str(pd.Timestamp.now().date()),source=SRC),
             summary=summary,segments=segments,linkage=linkage)

def _nan(v):
    if isinstance(v,float) and math.isnan(v): return None
    if isinstance(v,dict): return {k:_nan(x) for k,x in v.items()}
    if isinstance(v,list): return [_nan(x) for x in v]
    if isinstance(v,(np.floating,)): 
        return None if math.isnan(float(v)) else float(v)
    if isinstance(v,(np.integer,)): return int(v)
    return v
payload=_nan(payload)

with open(r'c:\Users\Lenovo\WorkBuddy\20261021112446\site_data.json','w',encoding='utf-8') as f:
    json.dump(payload,f,ensure_ascii=False)

print('REFRESH source_used:',SRC,'success:',REFRESH_STATUS['success'])
print('core profits:',len(core),[d['name'] for d in core])
print('strongest:',strong['name'],strong['hist_q'],'| latest',strong['latest'])
print('weakest:',weak['name'],weak['hist_q'],'| latest',weak['latest'])
print('movers_up(by 分位变化):',[(d['name'],d['chg20_q']) for d in movers_up])
print('movers_down(by 分位变化):',[(d['name'],d['chg20_q']) for d in movers_down])
print('asof_daily',asof_daily,'asof_weekly',asof_weekly)
print('EXCL counts:',{k[1]:len(v) for k,v in EXCL.items()})
print('oct_window items:',len(oct_window))
print('linkage pairs:',len(linkage),'| groups:',{'structural':sum(1 for l in linkage if l['group']=='structural'),'nr':sum(1 for l in linkage if l['group']=='nr')})
print('market_divergence items:',len(market_divergence))
print('segments indicators:',{k:len(v['indicators']) for k,v in segments.items()})
print('payload size KB:',round(len(json.dumps(payload,ensure_ascii=False))/1024,1))
print('OK')
