import urllib.request, json, time, sys

sys.stdout.reconfigure(encoding="utf-8")
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/vnd.github+json"}
base = "https://api.github.com/repos/hupeini7-crypto/rubber-profit-dashboard"


def get(u):
    return json.loads(urllib.request.urlopen(urllib.request.Request(u, headers=H), timeout=15).read().decode())


for i in range(12):
    run = get(base + "/actions/runs?per_page=1")["workflow_runs"][0]
    print(f"[{i+1}] {run['status']} | {run['conclusion']} | {run['head_sha'][:7]}")
    if run["status"] == "completed":
        break
    time.sleep(12)

time.sleep(8)
for i in range(6):
    try:
        r = urllib.request.urlopen(urllib.request.Request(
            "https://hupeini7-crypto.github.io/rubber-profit-dashboard/%E4%B8%8A%E4%B8%8B%E6%B8%B8%E5%88%A9%E6%B6%A6_%E5%B9%B4%E5%BA%A6%E5%AD%A3%E8%8A%82%E6%80%A7%E5%9B%BE%E8%A1%A8.html",
            headers={"User-Agent": "Mozilla/5.0"}), timeout=20)
        body = r.read()
        txt = body.decode("utf-8", "ignore")
        print("PAGE ->", r.status, "| bytes:", len(body), "| has 2026-11-10:", "2026-11-10" in txt)
        break
    except Exception as e:
        print("PAGE ->", type(e).__name__, "| retry")
        time.sleep(15)
