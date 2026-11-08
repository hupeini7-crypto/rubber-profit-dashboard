import subprocess, sys, os, time, urllib.request
from urllib.parse import quote
d = r'c:\Users\Lenovo\WorkBuddy\20261021112446'
PORT = 8910
flags = subprocess.CREATE_NEW_PROCESS_GROUP
if hasattr(subprocess, 'DETACHED_PROCESS'):
    flags |= subprocess.DETACHED_PROCESS
p = subprocess.Popen([sys.executable, '-m', 'http.server', str(PORT), '--directory', d],
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)
time.sleep(3)
base = f'http://127.0.0.1:{PORT}/'
try:
    files = urllib.request.urlopen(base, timeout=5).read().decode('utf-8', 'ignore')
    import re
    print('ROOT contains our html?', '上下游利润' in files)
except Exception as e:
    print('ROOT ERR', e)
f = '上下游利润_年度季节性图表.html'
url = base + quote(f)
try:
    r = urllib.request.urlopen(url, timeout=5)
    print('HTML STATUS', r.status, 'BYTES', len(r.read()))
except Exception as e:
    print('HTML ERR', e)
print('DONE')
