from __future__ import annotations

"""工具二：撰写 / 生成文档（compose）。

设计取舍：大模型负责「写什么」，工具负责「怎么排版」。
因此本工具接收已经写好的内容（Markdown 大纲/正文），只做确定性的 DOCX 排版输出，
不做任何自由发挥——这样生成的文档结构稳定、可复现，也符合计划书的分层原则。
"""

import json
import time

from pydantic import BaseModel, Field

from core import config
from core.ctx import get_session_id
from core.logger import log_call
from core.models import CODE_INTERNAL_ERROR, CODE_INVALID_INPUT, ToolResult
from core.safety import build_output_path
from core.tools import _docxkit as dk
from core.tools._mdrender import render_markdown_to_docx

TOOL_NAME = "compose_document"
TOOL_DESC = (
    "根据给定的 Markdown 内容新建一份 DOCX 文档并保存为新文件。"
    "当你需要「写一份报告/方案/总结/会议纪要」时，你自己在脑子里把正文写好，"
    "按 Markdown 组织好，再调用本工具落盘。支持 # 标题、- 列表、| 表格 |、> 引用。"
    "必须提供 outline 且不为空。"
)


class ComposeArgs(BaseModel):
    topic: str = Field(description="文档主题，用于命名输出文件，例如「项目周报」")
    outline: str = Field(
        description="Markdown 格式的正文内容。用 # / ## 表示标题，- 表示项目符号，| a | b | 表示表格。必须非空。"
    )
    title: str = Field(default="", description="文档大标题，留空则取 topic")
    author: str = Field(default="", description="作者名，会写在封面（可选）")
    with_cover: bool = Field(default=False, description="是否生成封面页")
    font: str = Field(default="", description="正文字体，留空用默认中文字体")
    font_size_pt: float = Field(default=config.DEFAULT_BODY_SIZE_PT, description="正文字号（磅）")
    line_spacing: float = Field(default=config.DEFAULT_LINE_SPACING, description="正文行距倍数")
    center_headings: bool = Field(default=True, description="一级标题是否居中")


def run(
    topic: str,
    outline: str,
    title: str = "",
    author: str = "",
    with_cover: bool = False,
    font: str = "",
    font_size_pt: float = config.DEFAULT_BODY_SIZE_PT,
    line_spacing: float = config.DEFAULT_LINE_SPACING,
    center_headings: bool = True,
    session_id: str = "",
) -> ToolResult:
    session_id = session_id or get_session_id()
    t0 = time.perf_counter()
    tool, action = "compose", "create"

    # 边界校验：任何一项缺失都如实返回失败，绝不生成空白文档充数
    if not str(topic or "").strip():
        r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "缺少文档主题（topic），无法确定要写什么",
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, {"topic": topic}, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
        return r
    if not str(outline or "").strip():
        r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "正文内容（outline）为空，无法生成文档",
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, {"topic": topic}, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
        return r
    if not isinstance(font_size_pt, (int, float)) or not (5 <= float(font_size_pt) <= 72):
        r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, f"字号非法：{font_size_pt}（应在 5~72 磅之间）",
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, {"font_size_pt": font_size_pt}, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
        return r

    out_path = build_output_path(None, action, ".docx", hint=topic)
    args = {"topic": topic, "title": title, "with_cover": with_cover, "font_size_pt": font_size_pt}

    try:
        render_markdown_to_docx(
            md_text=outline,
            title=title or topic,
            output_path=out_path,
            body_font=font or None,
            body_size_pt=float(font_size_pt),
            line_spacing=float(line_spacing),
            center_headings=bool(center_headings),
            add_cover=bool(with_cover),
            author=author,
        )
        doc = dk.open_docx(out_path)
        filled = dk.non_empty_paragraphs(doc)
        r = ToolResult.ok_result(
            tool, action,
            f"已生成《{title or topic}》：{len(filled)} 个非空段落、{len(doc.tables)} 个表格，约 {dk.word_count(doc)} 字",
            output_file=str(out_path),
            data={
                "title": title or topic,
                "paragraph_nonempty": len(filled),
                "table_count": len(doc.tables),
                "word_count": dk.word_count(doc),
                "outline": dk.document_outline(doc, 20),
            },
            elapsed_ms=int((time.perf_counter() - t0) * 1000),
        )
    except Exception as e:
        msg = f"生成文档失败：{type(e).__name__}: {e}"
        r = ToolResult.fail(tool, action, CODE_INTERNAL_ERROR, msg, elapsed_ms=int((time.perf_counter() - t0) * 1000))

    log_call(session_id, "tool", TOOL_NAME, action, args, r.ok, r.code, r.message,
             output_file=r.output_file, elapsed_ms=r.elapsed_ms)
    return r


def _lc_entry(
    topic: str,
    outline: str,
    title: str = "",
    author: str = "",
    with_cover: bool = False,
    font: str = "",
    font_size_pt: float = config.DEFAULT_BODY_SIZE_PT,
    line_spacing: float = config.DEFAULT_LINE_SPACING,
    center_headings: bool = True,
) -> str:
    return json.dumps(run(topic, outline, title, author, with_cover, font, font_size_pt, line_spacing, center_headings).to_dict(),
                      ensure_ascii=False)


def build_tool():
    from langchain_core.tools import StructuredTool
    return StructuredTool.from_function(
        func=_lc_entry, name=TOOL_NAME, description=TOOL_DESC, args_schema=ComposeArgs,
    )
