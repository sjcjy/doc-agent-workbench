from __future__ import annotations

"""调用记录：把每一次工具调用与大模型调用追加写入 JSONL。

计划书要求「显示调用记录」「多层按次调用工具」，这里统一落地：
一行 JSON = 一次调用，Streamlit 用表格展示，FastAPI 用 JSON 返回给小程序。
"""

import json
import time
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

from core import config

_LOCK = threading.Lock()


class _Timer:
    """简易计时器，用于记录工具耗时。"""

    def __init__(self) -> None:
        self.t0 = time.perf_counter()

    @property
    def ms(self) -> int:
        return int((time.perf_counter() - self.t0) * 1000)


def start_timer() -> _Timer:
    return _Timer()


def _jsonable(obj: Any, limit: int = 1200) -> Any:
    """把任意参数整理成可 JSON 序列化的形式，超长字符串截断，避免日志膨胀。"""
    if obj is None or isinstance(obj, (bool, int, float)):
        return obj
    if isinstance(obj, str):
        return obj[:limit] + ("…(截断)" if len(obj) > limit else "")
    if isinstance(obj, Path):
        return str(obj)
    if isinstance(obj, dict):
        return {str(k): _jsonable(v, limit) for k, v in list(obj.items())[:40]}
    if isinstance(obj, (list, tuple, set)):
        return [_jsonable(v, limit) for v in list(obj)[:40]]
    return str(obj)[:limit]


def log_call(
    session_id: str,
    layer: str,          # agent / tool / model / api / safety
    name: str,           # 工具名或模型名
    action: str = "",
    args: dict | None = None,
    ok: bool = True,
    code: str = "OK",
    message: str = "",
    output_file: str | Path | None = None,
    elapsed_ms: int = 0,
    backend: str = "",
    extra: dict | None = None,
) -> dict:
    """写一条调用记录到 logs/calls.jsonl，同时返回这条记录本身便于界面即时展示。"""
    record = {
        "ts": datetime.now().isoformat(timespec="milliseconds"),
        "session_id": session_id or "-",
        "layer": layer,
        "name": name,
        "action": action,
        "backend": backend,
        "ok": ok,
        "code": code,
        "message": (message or "")[:400],
        "args": _jsonable(args or {}),
        "output_file": str(output_file) if output_file else "",
        "elapsed_ms": elapsed_ms,
    }
    if extra:
        record["extra"] = _jsonable(extra)

    try:
        config.CALL_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _LOCK:
            with open(config.CALL_LOG_PATH, "a", encoding="utf-8") as f:
                f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        pass  # 日志失败绝不能拖垮主流程
    return record


def read_calls(session_id: str | None = None, limit: int = 300) -> list[dict]:
    """读取调用记录，可按会话过滤。按时间倒序返回最新的若干条。"""
    path: Path = config.CALL_LOG_PATH
    if not path.exists():
        return []
    rows: list[dict] = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    r = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if session_id and r.get("session_id") != session_id:
                    continue
                rows.append(r)
    except OSError:
        return []
    rows.reverse()
    return rows[:limit]


def new_session_id() -> str:
    return datetime.now().strftime("%Y%m%d%H%M%S") + "-" + str(int(time.time() * 1000))[-5:]


def clear_log() -> None:
    try:
        config.CALL_LOG_PATH.unlink(missing_ok=True)
    except OSError:
        pass
