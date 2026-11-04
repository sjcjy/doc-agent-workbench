# -*- coding: utf-8 -*-
"""端到端跑一遍「真实运行时」：走 HTTP 上传 → 工具调用 → 下载，看后端与产物。"""
import json
import mimetypes
import urllib.request
import uuid
from pathlib import Path

BASE = "http://127.0.0.1:8000"
SRC = Path(r"E:\Vibecoding\doc-agent-workbench\data\uploads\003重邮人工智能学院2027年初复试专业课参考书目+考试大纲.pdf")


def opener():
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def post_json(path, payload):
    req = urllib.request.Request(
        BASE + path, data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json"}, method="POST")
    with opener().open(req, timeout=180) as r:
        return json.loads(r.read().decode())


def post_file(path, field, filepath):
    boundary = "----wb" + uuid.uuid4().hex
    body = bytearray()
    body += f"--{boundary}\r\n".encode()
    body += f'Content-Disposition: form-data; name="{field}"; filename="{filepath.name}"\r\n'.encode()
    body += b"Content-Type: application/pdf\r\n\r\n"
    body += filepath.read_bytes() + b"\r\n"
    body += f"--{boundary}--\r\n".encode()
    req = urllib.request.Request(
        BASE + path, data=bytes(body),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}, method="POST")
    with opener().open(req, timeout=180) as r:
        return json.loads(r.read().decode())


def main():
    up = post_file("/api/upload", "file", SRC)
    print("upload:", json.dumps(up, ensure_ascii=False))

    res = post_json("/api/tool", {
        "tool": "convert_document",
        "args": {"path": up["path"], "target": "docx"},
    })
    print("tool ok:", res.get("ok"), res.get("code"))
    print("message:", res.get("message"))
    print("output:", res.get("output_file"))
    print("data:", json.dumps(res.get("data"), ensure_ascii=False)[:500])

    if res.get("output_file"):
        url = BASE + res["output_file"]
        with opener().open(url, timeout=120) as r:
            blob = r.read()
        out = Path(r"data\outputs\_runtime_download.docx")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(blob)
        print("downloaded:", out, len(blob), "bytes")

    logs = json.loads(opener().open(BASE + "/api/logs?limit=5", timeout=30).read().decode())
    for rec in logs.get("records", [])[-3:]:
        print("log:", rec.get("ts"), rec.get("name"), rec.get("ok"), rec.get("message")[:120])


if __name__ == "__main__":
    main()
