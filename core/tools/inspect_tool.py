from __future__ import annotations

"""工具一：检查文档（inspect）。

对应计划书「先检查后修改」原则——任何修改类任务都应该先跑一次 inspect，
拿到段落序号与标题结构，再决定改哪儿，避免盲改。
"""

import json
import time
from collections import Counter
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field

from core import config
from core.ctx import get_session_id
from core.logger import log_call
from core.models import CODE_INTERNAL_ERROR, ToolResult
from core.safety import InputGuard, sha256_of, validate_input_file
from core.tools import _docxkit as dk

TOOL_NAME = "inspect_document"
TOOL_DESC = (
    "检查一份 DOCX 文档的结构与内容，返回段落数、表格数、字数、标题大纲、"
    "段落序号、样式分布和开头预览。任何修改、调格式、转换任务之前都应先用它看一眼文档结构。"
    "当任务提到「这份文档」但不清楚具体内容时使用。"
)


class InspectArgs(BaseModel):
    path: str = Field(description="要检查的 DOCX 文件的绝对路径")
    keyword: str = Field(default="", description="可选：只在文档中检索该关键词，返回命中的段落序号")
    max_outline: int = Field(default=40, description="大纲最多返回多少条")


def run(path: str, keyword: str = "", max_outline: int = 40, session_id: str = "") -> ToolResult:
    session_id = session_id or get_session_id()
    t0 = time.perf_counter()
    tool, action = "inspect", "overview"

    ok_path, err = validate_input_file(path, tool, action)
    if err is not None:
        err.elapsed_ms = int((time.perf_counter() - t0) * 1000)
        log_call(session_id, "tool", TOOL_NAME, action, {"path": path}, False, err.code, err.message, elapsed_ms=err.elapsed_ms)
        return err
    src = ok_path

    with InputGuard(src) as guard:
        try:
            doc = dk.open_docx(src)
            paragraphs = doc.paragraphs
            total = len(paragraphs)
            filled = [(i, p) for i, p in enumerate(paragraphs) if dk.para_text(p).strip()]
            styles = Counter((p.style.name if p.style is not None else "Normal") for p in paragraphs if dk.para_text(p).strip())
            outline = dk.document_outline(doc, max_outline)
            data: dict[str, Any] = {
                "file": str(src),
                "file_name": src.name,
                "size_bytes": src.stat().st_size,
                "sha256_short": sha256_of(src)[:12],
                "paragraph_total": total,
                "paragraph_nonempty": len(filled),
                "table_count": len(doc.tables),
                "word_count": dk.word_count(doc),
                "style_distribution": dict(styles.most_common(10)),
                "outline": outline,
                "preview": dk.preview_text(doc),
            }

            action = "locate" if keyword else "overview"
            if keyword:
                hits = []
                for i, p in enumerate(paragraphs):
                    txt = dk.para_text(p)
                    if keyword in txt:
                        hits.append({"index": i, "text": txt.strip()[:120], "style": p.style.name})
                for ti, tb in enumerate(doc.tables):
                    for ri, row in enumerate(tb.rows):
                        for ci, cell in enumerate(row.cells):
                            if keyword in cell.text:
                                hits.append({"table": ti, "row": ri, "col": ci, "text": cell.text.strip()[:120]})
                data["keyword"] = keyword
                data["hits"] = hits
                data["hit_count"] = len(hits)
                msg = f"检索「{keyword}」命中 {len(hits)} 处" if hits else f"未在文档中找到「{keyword}」"
                data["found"] = bool(hits)
            else:
                msg = f"文档结构：{total} 段（非空 {len(filled)}）、{len(doc.tables)} 个表格、约 {data['word_count']} 字"

            elapsed = int((time.perf_counter() - t0) * 1000)
            res = ToolResult.ok_result(tool, action, msg, data=data, output_file=None, elapsed_ms=elapsed)
            return res
        except Exception as e:
            msg = f"文档解析失败：{type(e).__name__}: {e}"
            log_call(session_id, "tool", TOOL_NAME, action, {"path": path}, False, CODE_INTERNAL_ERROR, msg,
                     elapsed_ms=int((time.perf_counter() - t0) * 1000))
            return ToolResult.fail(tool, action, CODE_INTERNAL_ERROR, msg)

    # guard 退出后才有 after-hash，此时才能判定原文件是否被改动
    if guard.violated:
        return ToolResult.fail(tool, action, "SAFETY_BLOCKED", "原文件在检查过程中被修改，操作已中止",
                               elapsed_ms=int((time.perf_counter() - t0) * 1000))
    log_call(session_id, "tool", TOOL_NAME, res.action, {"path": path, "keyword": keyword},
             True, "OK", res.message, elapsed_ms=res.elapsed_ms, extra={"guard": guard.report()})
    return res


# ---------------------------------------------------------------- LangChain 包装
def as_json_result(res: ToolResult) -> str:
    """把 ToolResult 交给大模型时统一序列化为 JSON 字符串。"""
    return json.dumps(res.to_dict(), ensure_ascii=False)


def _lc_entry(path: str, keyword: str = "", max_outline: int = 40) -> str:
    return as_json_result(run(path=path, keyword=keyword, max_outline=max_outline))


def build_tool():
    from langchain_core.tools import StructuredTool
    return StructuredTool.from_function(
        func=_lc_entry,
        name=TOOL_NAME,
        description=TOOL_DESC,
        args_schema=InspectArgs,
    )
