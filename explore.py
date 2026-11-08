# -*- coding: utf-8 -*-
import os, sys, json, traceback

SRC = r'D:\AAA东方嘉信\备份数据\上下游利润.xlsx'
SRC_DIR = os.path.dirname(SRC)

print('=== win32com / Excel 可用性 ===')
try:
    import win32com.client
    print('win32com: AVAILABLE')
    try:
        xl = win32com.client.Dispatch('Excel.Application')
        print('Excel.Application: OK, version=', xl.Version)
        xl.Quit()
    except Exception as e:
        print('Excel.Application dispatch FAIL:', repr(e))
except Exception as e:
    print('win32com: NOT available ->', repr(e))

print('\n=== openpyxl 读取外部链接 ===')
import openpyxl
wb = openpyxl.load_workbook(SRC, data_only=False, read_only=True)
print('sheetnames:', wb.sheetnames)
# 外部链接
try:
    ext = wb._external_links
    print('external_links count:', len(ext) if ext else 0)
    for i, l in enumerate(ext or []):
        try:
            print('  link', i, '->', l.file_link.Target)
        except Exception as e:
            print('  link', i, 'attr err', e)
except Exception as e:
    print('external_links read err:', e)
wb.close()

print('\n=== 全部列名(规范去后缀前) ===')
import pandas as pd
xl = pd.ExcelFile(SRC)
for s in xl.sheet_names:
    df = pd.read_excel(xl, s, nrows=1)
    cols = list(df.columns)
    print(f'--- {s} ({len(cols)} cols) ---')
    for c in cols:
        print('   ', repr(c))

print('\n=== 搜索 泰混通关 相关列 ===')
for s in xl.sheet_names:
    df = pd.read_excel(xl, s, nrows=1)
    for c in df.columns:
        if '泰混' in str(c) or '通关' in str(c):
            print('FOUND in', s, '->', repr(c))

print('\n=== 半钢产能利用率 单元格公式 + 2015-06-30 值及相邻 ===')
# 找下游 sheet 中 半钢轮胎：产能利用率：中国（周） 列
wb2 = openpyxl.load_workbook(SRC, data_only=False)
ws = wb2['下游']
# header row
hdr = {}
for ci, cell in enumerate(ws[1], start=1):
    hdr[ci] = cell.value
target_col = None
for ci, v in hdr.items():
    if v and '半钢轮胎：产能利用率：中国' in str(v):
        target_col = ci
        print('半钢开工率 列号=', ci, '表头=', v)
        break
if target_col:
    # 找日期列(第1列)
    rows = []
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 1).value
        v = ws.cell(r, target_col).value
        rows.append((r, d, v))
    # 找 2015-06-30
    import datetime
    found_idx = None
    for i, (r, d, v) in enumerate(rows):
        try:
            if d and pd.to_datetime(d).date() == pd.Timestamp('2015-06-30').date():
                found_idx = i
                break
        except Exception:
            pass
    if found_idx is not None:
        print('2015-06-30 单元格公式:', repr(ws.cell(rows[found_idx][0], target_col).value))
        print('相邻±2行:')
        for j in range(max(0, found_idx-2), min(len(rows), found_idx+3)):
            print('   row', rows[j][0], 'date', rows[j][1], 'value', rows[j][2])
    else:
        print('2015-06-30 未在下游日期列找到')
    # 统计该列最大值
    vals = [v for (r, d, v) in rows if isinstance(v, (int, float))]
    if vals:
        print('该列 max=', max(vals), 'min=', min(vals), 'n=', len(vals), '异常>500计数=', sum(1 for v in vals if v > 500))
    else:
        print('该列无数值 (target_col 可能未命中，target_col=', target_col, ')')
wb2.close()

print('\n=== 外部依赖文件 国内需求(周度数据) 是否存在 ===')
for cand in ['国内需求(周度数据).xlsx', '国内需求(周度数据).xls']:
    p = os.path.join(SRC_DIR, cand)
    print(cand, '->', 'EXISTS' if os.path.exists(p) else 'NOT FOUND', p)
