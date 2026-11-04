# -*- coding: utf-8 -*-
"""API 层冒烟：验证 HTTP 接口可用（后续小程序对接的就是这些接口）。"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import httpx

from core import config

BASE = "http://127.0.0.1:8000"
SRC = str(config.SAMPLE_DIR / "样例1_项目周报.docx")


def pp(label, resp):
    try:
        body = resp.json()
    except Exception:
        body = resp.text[:200]
    print(f"[{resp.status_code}] {label}: {json.dumps(body, ensure_ascii=False)[:260]}")


def main():
    c = httpx.Client(timeout=120)

    pp("health", c.get(f"{BASE}/api/health"))

    r = c.get(f"{BASE}/api/files")
    print(f"[{r.status_code}] files: 共 {len(r.json()['files'])} 个")

    pp("task/查看结构", c.post(f"{BASE}/api/task", json={
        "query": "查看这份文档的结构", "file_path": SRC, "backend": "rule"}))

    r = c.post(f"{BASE}/api/task", json={
        "query": "把所有标题居中", "file_path": SRC, "backend": "rule"})
    d = r.json()
    print(f"[{r.status_code}] task/标题居中: ok={d.get('ok')} backend={d.get('backend')} "
          f"steps={len(d.get('steps', []))} outputs={len(d.get('outputs', []))}")

    if d.get("outputs"):
        name = Path(d["outputs"][0]).name
        rr = c.get(f"{BASE}/api/download/{name}")
        print(f"[{rr.status_code}] download/{name[:28]}: {len(rr.content)} bytes")

    pp("plan/转PDF", c.post(f"{BASE}/api/plan", json={
        "query": "转成 PDF", "file_path": SRC}))

    pp("tool/line_spacing", c.post(f"{BASE}/api/tool", json={
        "tool": "format_document",
        "args": {"path": SRC, "action": "line_spacing", "line_spacing": 2.0}}))

    pp("logs", c.get(f"{BASE}/api/logs?limit=3"))

    print("\nOpenAPI 文档： http://127.0.0.1:8000/docs")


if __name__ == "__main__":
    main()
