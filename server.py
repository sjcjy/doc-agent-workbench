# -*- coding: utf-8 -*-
"""文档工作台智能体 · HTTP 接口层。

存在的意义：把 core/ 的能力原样暴露成 HTTP API，
后续微信小程序直接对接这些接口即可，不需要重写任何业务逻辑。

启动：uvicorn server:app --reload --port 8000
接口一览：见 /docs 或 README.md
"""

import os
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

from core import config
from core.agent import execute, plan_only
from core.ctx import set_session_id
from core.logger import new_session_id, read_calls
from core.safety import build_output_path, save_upload, sha256_of, validate_input_file
from core.tools import convert_tool, dispatch
from core.tools import tool_descriptions

app = FastAPI(
    title="文档工作台智能体 API",
    description="供 Web 与微信小程序共用的后端。撰写 / 修改 / 格式 / 转换四项能力。",
    version="1.0.0",
)

# 小程序调试与本地前端需要跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 可选令牌鉴权：默认关闭。部署到公网/小程序时，在 .env 设 ACCESS_TOKEN 即生效。
# 小程序端在每个请求头带 Authorization: Bearer <token>。
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "").strip()


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    if ACCESS_TOKEN:
        path = request.url.path
        if path in ("/api/health", "/docs", "/openapi.json", "/redoc"):
            return await call_next(request)
        token = request.headers.get("Authorization", "").removeprefix("Bearer ").strip()
        if token != ACCESS_TOKEN:
            return JSONResponse(status_code=401, content={"detail": "unauthorized"})
    return await call_next(request)


# ---------------------------------------------------------------- 数据模型
class TaskRequest(BaseModel):
    query: str = Field(description="自然语言任务")
    file_path: str = Field(default="", description="当前文档路径")
    session_id: str = Field(default="", description="会话 ID，用于串联调用记录")
    backend: str = Field(default="", description="auto / llm / rule，留空取全局配置")


class ToolRequest(BaseModel):
    tool: str = Field(description="工具名")
    args: dict[str, Any] = Field(default_factory=dict, description="工具参数")
    session_id: str = Field(default="", description="会话 ID")


def _normalize_name(name: str) -> str:
    """大小写 + 全角/半角归一化，用于文件名容错匹配。"""
    name = Path(name).name
    s = name.lower()
    # 全角转半角
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def _to_download_url(abs_path: str) -> str:
    """把本地绝对路径转成可下载的相对 URL（供前端/小程序用）。"""
    if not abs_path:
        return ""
    name = Path(abs_path).name
    return f"/api/download/{quote(name)}"


def _abs_file(filename: str) -> Path:
    """按文件名定位输出文件，防越界；兼容中文/编码/大小写差异。"""
    candidates = [filename, unquote(filename)]
    for cand in candidates:
        name = Path(cand).name
        for root in (config.OUTPUT_DIR, config.UPLOAD_DIR, config.SAMPLE_DIR):
            p = root / name
            if p.exists():
                return p
    # 兜底：归一化模糊匹配（应对隧道对中文路径的编码差异）
    target = _normalize_name(filename)
    for root in (config.OUTPUT_DIR, config.UPLOAD_DIR, config.SAMPLE_DIR):
        if not root.exists():
            continue
        for p in root.glob("*"):
            if _normalize_name(p.name) == target:
                return p
    raise HTTPException(status_code=404, detail="文件不存在")


# ---------------------------------------------------------------- 基础
@app.get("/")
def index() -> FileResponse:
    """轻量 Web 前端（单文件），穿透 8000 端口即可在公网稳定使用。"""
    hp = ROOT / "web" / "index.html"
    if not hp.exists():
        raise HTTPException(status_code=404, detail="前端文件缺失：web/index.html")
    return FileResponse(str(hp), media_type="text/html")


@app.get("/api/health")
def health() -> dict:
    return {
        "status": "ok",
        "version": app.version,
        "has_model_key": config.has_model_key(),
        "model_provider": config.MODEL_PROVIDER,
        "model": config.ZHIPU_MODEL if config.MODEL_PROVIDER.startswith("zhipu") else config.OPENAI_MODEL,
        "agent_backend": config.AGENT_BACKEND,
        "pdf_backends": convert_tool.pdf_backends(),
    }


@app.get("/api/tools")
def tools_list() -> dict:
    return {"tools": tool_descriptions()}


# ---------------------------------------------------------------- 文件
@app.get("/api/files")
def list_files() -> dict:
    """列出可选文档：上传 / 样例 / 最近输出。"""
    out: list[dict] = []
    for kind, root in (("sample", config.SAMPLE_DIR), ("upload", config.UPLOAD_DIR), ("output", config.OUTPUT_DIR)):
        if not root.exists():
            continue
        files = sorted(root.glob("*"), key=lambda x: x.stat().st_mtime, reverse=True)
        for p in files[:50]:
            if p.suffix.lower() not in (config.SUPPORTED_UPLOAD_EXT | {".pdf"}):
                continue
            out.append({
                "name": p.name,
                "path": str(p),
                "kind": kind,
                "size": p.stat().st_size,
                "modified": int(p.stat().st_mtime),
            })
    return {"files": out}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)) -> dict:
    data = await file.read()
    target, err = save_upload(data, file.filename or "")
    if err is not None:
        raise HTTPException(status_code=400, detail={"code": err.code, "message": err.message})
    return {
        "ok": True,
        "path": str(target),
        "name": target.name,
        "size": target.stat().st_size,
        "sha256": sha256_of(target)[:16],
    }


@app.get("/api/download/{filename}")
def download(filename: str) -> FileResponse:
    p = _abs_file(filename)
    return FileResponse(str(p), filename=p.name, media_type="application/octet-stream")


@app.post("/api/hash")
def file_hash(payload: dict) -> dict:
    """计算文件哈希，用于「原文件未被改动」的验收核对。"""
    path = str(payload.get("path", "")).strip()
    okp, err = validate_input_file(path, "api", "hash")
    if err is not None:
        raise HTTPException(status_code=400, detail={"code": err.code, "message": err.message})
    return {"path": path, "sha256": sha256_of(okp)}


# ---------------------------------------------------------------- 核心：任务
@app.post("/api/task")
def run_task(req: TaskRequest) -> dict:
    """自然语言任务入口。Agent 自动决定调用哪些工具。"""
    if not req.query.strip():
        raise HTTPException(status_code=400, detail={"code": "INVALID_INPUT", "message": "任务内容为空"})
    sid = req.session_id or new_session_id()
    set_session_id(sid)
    result = execute(req.query.strip(), req.file_path, sid, backend=req.backend or None)
    result["session_id"] = sid
    # 把本地绝对路径改成可下载的相对 URL，前端/小程序才能拿到文件
    result["outputs"] = [_to_download_url(o) for o in (result.get("outputs") or [])]
    for s in (result.get("steps") or []):
        r = s.get("result") or {}
        if r.get("output_file"):
            r["output_file"] = _to_download_url(r["output_file"])
    return result


@app.post("/api/plan")
def plan_task(req: TaskRequest) -> dict:
    """只做意图识别不执行，返回将要调用的工具与参数（规则路由支持）。"""
    return {"query": req.query, "file_path": req.file_path, "plan": plan_only(req.query, req.file_path)}


@app.post("/api/tool")
def call_tool(req: ToolRequest) -> dict:
    """直接调用某个工具，绕过 Agent，适合小程序里做确定性的固定按钮。"""
    sid = req.session_id or new_session_id()
    set_session_id(sid)
    try:
        res = dispatch(req.tool, **req.args, session_id=sid)
    except KeyError as e:
        raise HTTPException(status_code=404, detail={"code": "UNKNOWN_TOOL", "message": str(e)})
    except Exception as e:
        raise HTTPException(status_code=400, detail={"code": "INVALID_ARGS", "message": f"{type(e).__name__}: {e}"})
    result = res.to_dict()
    result["session_id"] = sid
    if result.get("output_file"):
        result["output_file"] = _to_download_url(result["output_file"])
    # 拆分会产生多份文件，data["files"] 里也全是本地绝对路径，同样要转成可下载 URL，
    # 否则前端和小程序拿到的是本机路径，点了打不开。
    data = result.get("data") or {}
    if result.get("action") == "split" and isinstance(data.get("files"), list):
        data["files"] = [_to_download_url(f) for f in data["files"]]
    if not res.ok:
        # 工具失败仍返回 200，让前端能展示结构化错误；code 用于区分
        result["http_ok"] = False
    return result


# ---------------------------------------------------------------- 调用记录
@app.get("/api/logs")
def get_logs(session_id: str = "", limit: int = 200) -> dict:
    return {"records": read_calls(session_id or None, limit)}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)
