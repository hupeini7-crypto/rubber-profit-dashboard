# -*- coding: utf-8 -*-
"""
一键更新并发布：重建数据 → 渲染页面 → 一致性校验 → git 提交推送。
推送到 GitHub main 分支后，GitHub Actions 自动发布到 GitHub Pages。

用法：
    python deploy.py            # 完整流程：build → render → verify → commit+push
    python deploy.py --skip-build   # 跳过数据重建，仅推送现有页面（调试用）
"""
import subprocess
import sys
import os
import re
import shutil
import datetime

WS = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
MAIN_HTML = '上下游利润_年度季节性图表.html'

def run(cmd, desc, allow_fail=False):
    print(f'\n===== {desc} =====')
    r = subprocess.run(cmd, cwd=WS, shell=True, capture_output=True,
                       text=True, encoding='utf-8', errors='replace')
    out = ((r.stdout or '') + (r.stderr or '')).strip()
    print(out[-2000:] if out else '(no output)')
    if r.returncode != 0 and not allow_fail:
        sys.exit(f'[FAIL] {desc}（退出码 {r.returncode}），流程中止。')
    return r

def git_dirty():
    r = subprocess.run('git status --porcelain', cwd=WS, shell=True,
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    return [l for l in (r.stdout or '').splitlines() if l.strip()]

# 1) 数据重建（build_site.py 内部先尝试 COM 刷新外部链接）
if '--skip-build' not in sys.argv:
    run(f'"{PY}" build_site.py', 'STEP 1/4 数据重建 build_site.py')
    run(f'"{PY}" render_html.py', 'STEP 1/4 页面渲染 render_html.py')
else:
    print('跳过数据重建（--skip-build）')

# 2) 一致性校验：不一致则中止发布
r = run(f'"{PY}" verify_consistency.py', 'STEP 2/4 一致性校验 verify_consistency.py')
m = re.search(r'不一致项数:\s*(\d+)\s*/\s*(\d+)', r.stdout or '')
if not m:
    sys.exit('[FAIL] 无法解析校验结果，流程中止。')
mismatch, total = int(m.group(1)), int(m.group(2))
print(f'>>> 校验结果：{mismatch} 不一致 / {total} 项')
if mismatch > 0:
    sys.exit('[FAIL] 存在不一致项，中止发布。请先排查数据问题。')

# 3) 生成 index.html（GitHub Pages 入口，直接跳转主页面）
index_path = os.path.join(WS, 'index.html')
with open(index_path, 'w', encoding='utf-8') as f:
    f.write('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
            '<meta http-equiv="refresh" content="0;url=%s">'
            '<title>天然橡胶产业链利润分析</title></head>'
            '<body>正在跳转… <a href="%s">若未自动跳转请点此</a></body></html>'
            % (MAIN_HTML, MAIN_HTML))
print('\nSTEP 3/4 已生成 index.html（Pages 入口跳转）')

# 4) git 提交推送
print('\nSTEP 4/4 git 提交推送')
dirty = git_dirty()
if not dirty:
    print('无变更，跳过提交与推送（页面已是最新）。')
else:
    print(f'变更 {len(dirty)} 项：')
    for l in dirty[:20]:
        print('  ' + l)
    run('git add -A', 'git add')
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    run(f'git commit -m "auto: data update {ts}"', 'git commit')
    run('git push origin main', 'git push origin main')
    print('\n>>> 已推送。GitHub Actions 将自动发布到 Pages（约 1-2 分钟生效）。')

print('\n[DONE] 部署流程完成。')
