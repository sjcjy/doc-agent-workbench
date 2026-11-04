from __future__ import annotations

"""工具四：调整文档格式（format）。

覆盖计划书要求的四类操作：标题居中、字体字号、行距、首行缩进，
外加粗体/斜体/颜色/段后间距，并提供一个「一键排版」预设 preset_report。
每次操作都返回逐段变更明细，方便在 Word 里逐条核对（验收时可直接截图）。
"""

import json
import time

from pydantic import BaseModel, Field
from typing_extensions import Literal

from core import config
from core.ctx import get_session_id
from core.logger import log_call
from core.models import CODE_INTERNAL_ERROR, CODE_INVALID_INPUT, CODE_NO_CHANGE, ToolResult
from core.safety import InputGuard, build_output_path, validate_input_file
from core.tools import _docxkit as dk

TOOL_NAME = "format_document"
TOOL_DESC = (
    "调整 DOCX 的格式，输出为新文件。"
    "action=align 对齐（alignment 取 left/center/right/justify，常用于「标题居中」）；"
    "action=font 字体字号（font_name 中文字体、size_pt 字号、bold 加粗）；"
    "action=line_spacing 行距（line_spacing 倍率，如 1.5）；"
    "action=indent 首行缩进（indent_chars 字符数，通常 2）；"
    "action=preset_report 一键套用报告排版（标题居中、正文宋体小四、1.5 倍行距、首行缩进 2 字符）。"
    "target 决定作用范围：all / headings / body / first / 段落序号 / 段落中包含的文字。"
)

FormatAction = Literal["align", "font", "line_spacing", "indent", "bold", "preset_report"]


class FormatArgs(BaseModel):
    path: str = Field(description="源 DOCX 文件绝对路径（只读）")
    action: FormatAction = Field(description="格式动作")
    target: str = Field(default="all", description="作用范围：all/headings/body/first/段落序号/段落内含文字")
    alignment: str = Field(default="center", description="align 时的对齐方式")
    font_name: str = Field(default="", description="中文字体名，如 宋体 / 黑体 / 微软雅黑")
    en_font: str = Field(default="", description="西文字体名，如 Times New Roman")
    size_pt: float = Field(default=0, description="字号（磅），0 表示不改；常用 12=小四 14=四号 16=三号 18=小二")
    bold: bool | None = Field(default=None, description="是否加粗，None 表示不改")
    italic: bool | None = Field(default=None, description="是否斜体，None 表示不改")
    line_spacing: float = Field(default=1.5, description="行距倍数")
    line_spacing_exact: bool = Field(default=False, description="true 时 line_spacing 按磅值固定行距")
    indent_chars: float = Field(default=2.0, description="首行缩进字符数")
    indent_pt: float = Field(default=0, description="首行缩进磅值，优先于 indent_chars")
    space_after_pt: float = Field(default=0, description="段后间距（磅），0 表示不改")


def select_paragraphs(doc, target: str) -> list[tuple[int, object]]:
    """按 target 描述挑选目标段落。"""
    t = str(target or "all").strip()
    paras = doc.paragraphs
    if t in ("", "all", "全文", "全部"):
        return [(i, p) for i, p in enumerate(paras) if dk.para_text(p).strip()]
    if t in ("headings", "heading", "标题"):
        return [(i, p) for i, p in enumerate(paras) if dk.is_heading(p)]
    if t in ("body", "正文"):
        return [(i, p) for i, p in enumerate(paras) if not dk.is_heading(p) and dk.para_text(p).strip()]
    if t in ("first", "title", "首段", "标题段"):
        return [(0, paras[0])] if paras else []
    if t.lstrip("-").isdigit():
        idx = int(t)
        if -len(paras) <= idx < len(paras):
            return [(idx % len(paras), paras[idx % len(paras)])]
        return []
    return [(i, p) for i, p in enumerate(paras) if t in dk.para_text(p)]


def run(
    path: str,
    action: str,
    target: str = "all",
    alignment: str = "center",
    font_name: str = "",
    en_font: str = "",
    size_pt: float = 0,
    bold: bool | None = None,
    italic: bool | None = None,
    line_spacing: float = 1.5,
    line_spacing_exact: bool = False,
    indent_chars: float = 2.0,
    indent_pt: float = 0,
    space_after_pt: float = 0,
    session_id: str = "",
) -> ToolResult:
    session_id = session_id or get_session_id()
    t0 = time.perf_counter()
    tool = "format"
    valid_actions = ("align", "font", "line_spacing", "indent", "bold", "preset_report")
    args = {"path": path, "action": action, "target": target, "size_pt": size_pt,
            "font_name": font_name, "alignment": alignment, "line_spacing": line_spacing}

    ok_path, err = validate_input_file(path, tool, action)
    if err is not None:
        err.elapsed_ms = int((time.perf_counter() - t0) * 1000)
        log_call(session_id, "tool", TOOL_NAME, action, args, False, err.code, err.message, elapsed_ms=err.elapsed_ms)
        return err
    src = ok_path

    if action not in valid_actions:
        r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                            f"不支持的格式动作：{action}（可选：{'、'.join(valid_actions)}）",
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
        return r

    with InputGuard(src) as guard:
        try:
            doc = dk.open_docx(src)
            changes: list[dict] = []
            out_path = build_output_path(src, action, ".docx")
            size_arg = float(size_pt) if size_pt else None

            if action == "preset_report":
                # 一键排版：一级标题居中，正文宋体小四 + 1.5 倍行距 + 首行缩进 2 字符
                for i, p in enumerate(doc.paragraphs):
                    t = dk.para_text(p).strip()
                    if not t:
                        continue
                    if dk.is_heading(p):
                        if dk.heading_level(p) == 1:
                            dk.align_paragraph(p, "center")
                        dk.set_para_font(p, cn_font=font_name or "黑体", en_font=en_font or None, bold=True)
                        changes.append({"index": i, "ops": ["标题字体"]})
                    else:
                        dk.set_para_font(p, cn_font=font_name or config.DEFAULT_CN_FONT, en_font=en_font or None, size_pt=size_arg or 12)
                        dk.set_line_spacing(p, 1.5)
                        dk.set_first_line_indent(p, chars=2)
                        dk.set_space_after(p, 6 if space_after_pt == 0 else space_after_pt)
                        changes.append({"index": i, "ops": ["字体", "1.5倍行距", "首行缩进2字符"]})
                summary = f"已套用报告排版，影响 {len(changes)} 段"

            else:
                targets = select_paragraphs(doc, target)
                if not targets:
                    r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                                        f"作用范围 target={target!r} 未匹配到任何段落，请换一个序号或关键词",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r

                for i, p in targets:
                    ops: list[str] = []
                    if action == "align":
                        if dk.align_paragraph(p, alignment):
                            ops.append(f"对齐={alignment}")
                    elif action == "font":
                        before = (p.runs[0].font.name if p.runs else None, p.runs[0].font.size.pt if p.runs and p.runs[0].font.size else None)
                        dk.set_para_font(p, cn_font=font_name or None, en_font=en_font or None,
                                         size_pt=size_arg, bold=bold)
                        after = (p.runs[0].font.name if p.runs else None, p.runs[0].font.size.pt if p.runs and p.runs[0].font.size else None)
                        if before != after or size_arg:
                            ops.append(f"字体={font_name or en_font or 'inherit'}")
                            if size_arg:
                                ops.append(f"字号={size_arg}pt")
                            if bold is not None:
                                ops.append("加粗" if bold else "取消加粗")
                    elif action == "line_spacing":
                        if dk.set_line_spacing(p, line_spacing, exact=line_spacing_exact):
                            ops.append(f"行距={'固定' + str(line_spacing) + 'pt' if line_spacing_exact else str(line_spacing) + '倍'}")
                    elif action == "indent":
                        if dk.set_first_line_indent(p, chars=indent_chars if not indent_pt else None, pt=indent_pt or None):
                            ops.append(f"首行缩进={indent_pt or indent_chars}{'pt' if indent_pt else '字符'}")
                    elif action == "bold":
                        for run in (p.runs or [p.add_run("")]):
                            run.font.bold = True if bold is None else bold
                        ops.append("加粗" if (bold is None or bold) else "取消加粗")
                    if space_after_pt and action != "preset_report":
                        if dk.set_space_after(p, space_after_pt):
                            ops.append(f"段后={space_after_pt}pt")
                    if ops:
                        changes.append({"index": i, "type": action, "ops": ops,
                                        "text": dk.para_text(p).strip()[:60]})
                summary = f"{action} 已应用于 {len(changes)} 个段落" if changes else "未产生任何格式变化"

            if not changes:
                r = ToolResult.fail(tool, action, CODE_NO_CHANGE,
                                    f"{summary}：目标段落已经是该格式，或参数未触发任何变更",
                                    elapsed_ms=int((time.perf_counter() - t0) * 1000))
                log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                return r

            doc.save(str(out_path))
            r = ToolResult.ok_result(tool, action, f"{summary}，已另存为 {out_path.name}",
                                     output_file=str(out_path), changes=changes,
                                     data={"source": str(src), "affected": len(changes)},
                                     elapsed_ms=int((time.perf_counter() - t0) * 1000))
            return r

        except Exception as e:
            msg = f"调整格式时出错：{type(e).__name__}: {e}"
            r = ToolResult.fail(tool, action, CODE_INTERNAL_ERROR, msg, elapsed_ms=int((time.perf_counter() - t0) * 1000))
            log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
            return r

    if guard.violated:
        return ToolResult.fail(tool, action, "SAFETY_BLOCKED", "源文件在处理过程中被改动，已中止",
                               elapsed_ms=int((time.perf_counter() - t0) * 1000))
    log_call(session_id, "tool", TOOL_NAME, action, args, r.ok, r.code, r.message,
             output_file=r.output_file, elapsed_ms=r.elapsed_ms, extra={"guard": guard.report()})
    return r


def _lc_entry(path: str, action: str, target: str = "all", alignment: str = "center",
              font_name: str = "", en_font: str = "", size_pt: float = 0,
              bold: bool | None = None, italic: bool | None = None, line_spacing: float = 1.5,
              line_spacing_exact: bool = False, indent_chars: float = 2.0, indent_pt: float = 0,
              space_after_pt: float = 0) -> str:
    res = run(path, action, target, alignment, font_name, en_font, size_pt, bold, italic,
              line_spacing, line_spacing_exact, indent_chars, indent_pt, space_after_pt)
    return json.dumps(res.to_dict(), ensure_ascii=False)


def build_tool():
    from langchain_core.tools import StructuredTool
    return StructuredTool.from_function(
        func=_lc_entry, name=TOOL_NAME, description=TOOL_DESC, args_schema=FormatArgs,
    )
