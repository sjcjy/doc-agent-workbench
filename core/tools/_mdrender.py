from __future__ import annotations

"""轻量 Markdown → DOCX 渲染器。

被「撰写」和「MD/TXT→DOCX 转换」两个工具共用。
支持：1-4 级标题、无序/有序列表、引用、简单表格、加粗/行内代码、**分隔线。
不支持整块的复杂嵌套——遇到不认识的行就按正文段落渲染，保证输出永远能打开。
"""

import re
from typing import Iterable

from docx import Document
from docx.shared import Pt

from core import config
from core.tools import _docxkit as dk

_RE_H = re.compile(r"^(#{1,6})\s+(.*)$")
_RE_UL = re.compile(r"^[-*+]\s+(.*)$")
_RE_OL = re.compile(r"^\d+[.)]\s+(.*)$")
_RE_QUOTE = re.compile(r"^>\s?(.*)$")
_RE_HR = re.compile(r"^([-*_])\1{2,}$")
_RE_TABLE_ROW = re.compile(r"^\|(.+)\|\s*$")
_RE_TABLE_SEP = re.compile(r"^\|[\s:|-]+\|\s*$")
_RE_BOLD = re.compile(r"\*\*(.+?)\*\*")
_RE_CODE = re.compile(r"`([^`]+)`")

LIST_BULLET = ("List Bullet", "List Bullet 2", "项目符号", "项目符号 2")
LIST_NUMBER = ("List Number", "List Number 2", "编号列表", "编号列表 2")
QUOTE_STYLE = ("Quote", "Intense Quote", "引用", "明显引用")


def _pick(doc: Document, candidates: Iterable[str]) -> str | None:
    """在候选样式名里挑第一个文档中真实存在的样式。"""
    names = [c for c in candidates if c]
    if not names:
        return None
    return dk.resolve_style(doc, names[0], names[1:])


def _render_inline(doc: Document, p, text: str, size_pt: float | None = None, bold: bool = False) -> None:
    """处理 **粗体**、`代码` 与普通文本混排。"""
    pos = 0
    for m in re.finditer(r"\*\*(.+?)\*\*|`([^`]+)`", text):
        if m.start() > pos:
            r = p.add_run(text[pos:m.start()])
            dk.set_run_font(r, size_pt=size_pt, bold=bold)
        if m.group(1) is not None:
            r = p.add_run(m.group(1))
            dk.set_run_font(r, size_pt=size_pt, bold=True)
        else:
            r = p.add_run(m.group(2))
            dk.set_run_font(r, cn_font="Consolas", en_font="Consolas", size_pt=(size_pt or 12) - 1)
        pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:])
        dk.set_run_font(r, size_pt=size_pt, bold=bold)


def render_markdown_to_docx(
    md_text: str,
    title: str = "",
    output_path=None,
    body_font: str | None = None,
    body_size_pt: float = config.DEFAULT_BODY_SIZE_PT,
    line_spacing: float = config.DEFAULT_LINE_SPACING,
    first_line_indent_chars: float = 2.0,
    center_headings: bool = True,
    add_cover: bool = False,
    author: str = "",
) -> Document:
    """把 Markdown 文本渲染成一个 Document 对象（若给出 output_path 则同时保存）。

    返回的 Document 调用方可继续加工。
    """
    doc = Document()

    # 正文默认样式：中文字体 + 字号，避免生成出来是 Word 默认的等线/Calibri
    try:
        normal = doc.styles["Normal"]
        normal.font.name = body_font or config.DEFAULT_CN_FONT
        normal.font.size = Pt(body_size_pt)
        rpr = normal.element.get_or_add_rPr()
        ns = rpr.find(dk.qn("w:rFonts"))
        if ns is None:
            ns = rpr.makeelement(dk.qn("w:rFonts"), {})
            rpr.insert(0, ns)
        ns.set(dk.qn("w:eastAsia"), body_font or config.DEFAULT_CN_FONT)
        ns.set(dk.qn("w:ascii"), body_font or config.DEFAULT_CN_FONT)
        ns.set(dk.qn("w:hAnsi"), body_font or config.DEFAULT_CN_FONT)
    except Exception:
        pass

    if add_cover:
        cover_title = doc.add_paragraph()
        r = cover_title.add_run(title or "文档")
        dk.set_run_font(r, cn_font=body_font or config.DEFAULT_CN_FONT, size_pt=26, bold=True)
        cover_title.alignment = 1  # WD_ALIGN_PARAGRAPH.CENTER
        for _ in range(6):
            doc.add_paragraph()
        if author:
            sub = doc.add_paragraph()
            r2 = sub.add_run(f"编制：{author}")
            dk.set_run_font(r2, size_pt=14)
            sub.alignment = 1
        doc.add_page_break()

    if title and not add_cover:
        dk.add_heading(doc, title, 1)
        if center_headings:
            dk.align_paragraph(doc.paragraphs[-1], "center")

    lines = (md_text or "").replace("\r\n", "\n").split("\n")
    table_buf: list[list[str]] = []
    i = 0
    n = len(lines)
    while i < n:
        raw = lines[i].rstrip()
        line = raw.strip()

        if not line:
            i += 1
            continue

        # 表格行先攒起来，遇到非表格行再一次性渲染
        if _RE_TABLE_ROW.match(line) and not _RE_TABLE_SEP.match(line):
            cells = [c.strip() for c in line.strip("|").split("|")]
            table_buf.append(cells)
            i += 1
            continue

        if table_buf:
            _flush_table(doc, table_buf, body_size_pt)
            table_buf = []

        if _RE_HR.match(line) or _RE_TABLE_SEP.match(line):
            i += 1
            continue

        mh = _RE_H.match(line)
        if mh:
            level = min(len(mh.group(1)), 4)
            p = dk.add_heading(doc, mh.group(2).strip(), level)
            if level == 1 and center_headings:
                dk.align_paragraph(p, "center")
            i += 1
            continue

        if "@ @ @" in line:  # 换页标记
            doc.add_page_break()
            i += 1
            continue

        mq = _RE_QUOTE.match(line)
        if mq:
            style = _pick(doc, QUOTE_STYLE)
            p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
            _render_inline(doc, p, mq.group(1), size_pt=body_size_pt)
            i += 1
            continue

        mu = _RE_UL.match(line)
        if mu:
            style = _pick(doc, LIST_BULLET)
            p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
            if style is None:
                p.add_run("· ")
            _render_inline(doc, p, mu.group(1), size_pt=body_size_pt)
            i += 1
            continue

        mo = _RE_OL.match(line)
        if mo:
            style = _pick(doc, LIST_NUMBER)
            p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
            _render_inline(doc, p, mo.group(1), size_pt=body_size_pt)
            i += 1
            continue

        # 普通正文
        p = doc.add_paragraph()
        _render_inline(doc, p, line, size_pt=body_size_pt)
        dk.set_line_spacing(p, line_spacing)
        dk.set_first_line_indent(p, chars=first_line_indent_chars)
        i += 1

    if table_buf:
        _flush_table(doc, table_buf, body_size_pt)

    if output_path:
        doc.save(str(output_path))
    return doc


def _flush_table(doc: Document, rows: list[list[str]], size_pt: float) -> None:
    """渲染 Markdown 表格，第一行为表头。"""
    if not rows:
        return
    width = max(len(r) for r in rows)
    try:
        table = doc.add_table(rows=0, cols=width)
        table.style = dk.resolve_style(doc, "Table Grid", "网格型") or table.style
    except Exception:
        return
    for ri, row in enumerate(rows):
        cells = row + [""] * (width - len(row))
        cells_row = table.add_row()
        for ci, val in enumerate(cells[:width]):
            cell = cells_row.cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            _render_inline(doc, p, val, size_pt=size_pt, bold=(ri == 0))
            dk.set_line_spacing(p, 1.15)
