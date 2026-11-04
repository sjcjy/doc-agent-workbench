"""通过 HTTP 走一遍结构操作的四个动作，验证 API 层与下载链接都通。"""
import json
import sys
import urllib.parse
import urllib.request
from pathlib import Path

BASE = "http://127.0.0.1:8000"
OP = urllib.request.build_opener(urllib.request.ProxyHandler({}))

A = str((Path("samples/样例1_项目周报.docx").resolve()))
B = str((Path("samples/样例2_实验报告.docx").resolve()))


def call(tool, args):
    body = json.dumps({"tool": tool, "args": args}, ensure_ascii=False).encode()
    req = urllib.request.Request(f"{BASE}/api/tool", data=body,
                                 headers={"Content-Type": "application/json"})
    with OP.open(req, timeout=180) as r:
        return json.loads(r.read())


def check_download(url: str, tag: str):
    full = url if url.startswith("http") else BASE + url
    try:
        with OP.open(full, timeout=60) as r:
            data = r.read()
        print(f"   {tag} 下载 {r.status} {len(data)} bytes  {url[:60]}")
        return len(data) > 0
    except Exception as e:
        print(f"   {tag} 下载失败 {type(e).__name__}: {e}  {url[:80]}")
        return False


ok_all = True
for tag, args in [
    ("merge", {"path": A, "action": "merge", "other_paths": B}),
    ("split", {"path": A, "action": "split", "level": 1}),
    ("tables", {"path": A, "action": "extract_tables"}),
    ("diff", {"path": A, "action": "diff", "other_paths": B}),
]:
    r = call("structure_document", args)
    print(f"[{tag}] ok={r.get('ok')} code={r.get('code')} | {r.get('message','')[:70]}")
    if not r.get("ok"):
        ok_all = False
        continue
    if r.get("output_file"):
        ok_all &= check_download(r["output_file"], tag)
    files = (r.get("data") or {}).get("files") or []
    if args["action"] == "split":
        print(f"   拆分产出 {len(files)} 份，逐个下载：")
        for i, f in enumerate(files, 1):
            ok_all &= check_download(f, f"{tag}#{i}")

print("=" * 50)
print("全部通过" if ok_all else "有失败项，见上")
sys.exit(0 if ok_all else 1)
