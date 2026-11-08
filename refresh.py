# -*- coding: utf-8 -*-
import os, sys, time, traceback, json

SRC = r'D:\AAA东方嘉信\备份数据\上下游利润.xlsx'
OUT = r'c:\Users\Lenovo\WorkBuddy\20261021112446\上下游利润_refreshed.xlsx'

log = []
def L(m):
    log.append(m); print(m)

try:
    import win32com.client
    import pythoncom
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.EnableEvents = False
    L('Excel started, version=%s' % xl.Version)

    # 列出当前链接源
    wb = xl.Workbooks.Open(SRC, UpdateLinks=0)  # 0=不自动更新，手动控制
    try:
        links = wb.LinkSources(1)  # xlExcelLinks
        L('LinkSources(xlExcelLinks): %r' % (list(links) if links else None))
    except Exception as e:
        L('LinkSources err: %r' % e)
        links = []

    # 检查每个链接文件是否可访问
    for lnk in (links or []):
        p = lnk
        # 可能是 file:// 或绝对路径
        import urllib.parse
        if p.lower().startswith('file:'):
            p = urllib.parse.unquote(p[7:])
        exists = os.path.exists(p)
        L('  链接: %s -> %s' % (p, 'EXISTS' if exists else 'MISSING'))

    L('开始 RefreshAll ...')
    try:
        wb.RefreshAll()
        time.sleep(8)  # 给外部链接刷新留时间
        L('RefreshAll 调用完成(无异常)')
    except Exception as e:
        L('RefreshAll 异常: %r' % e)

    # 强制重新计算
    try:
        xl.CalculateFullRebuild()
        L('CalculateFullRebuild 完成')
    except Exception as e:
        L('Calculate err: %r' % e)

    # 保存副本（不覆盖原始，便于回滚与核查）
    try:
        wb.SaveAs(OUT, FileFormat=51)  # 51=xlsx
        L('已保存刷新副本: %s' % OUT)
    except Exception as e:
        L('SaveAs 异常: %r' % e)
    wb.Close(SaveChanges=False)
    xl.Quit()
    L('DONE')
except Exception as e:
    L('FATAL: %r' % e)
    traceback.print_exc()
finally:
    try:
        xl.Quit()
    except Exception:
        pass

open(r'c:\Users\Lenovo\WorkBuddy\20261021112446\refresh_log.txt','w',encoding='utf-8').write('\n'.join(log))
