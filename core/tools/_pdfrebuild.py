# -*- coding: utf-8 -*-
"""PDF → DOCX 的"结构重建"转换器。

为什么不复制坐标：Word 自带的 PDF Reflow 会把每一段文字做成浮动文本框（本文档实测
1256 个），位置靠锚点段落维持，字号/行高/页边距一改就整体错位，而且产物根本没法
编辑。pdf2docx 则把 PDF 绝对坐标写成段落缩进，同样一改就崩。

这里换个思路：不复制"看起来的位置"，而是从 PDF 里读出**结构**——
表格靠网格线（横线/竖线）还原成真正的 Word 表格，正文按行块还原成真正的段落，
页眉页脚放进节属性。产物是可编辑、可自动换行的 Word 文档。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

CJK_RE = re.compile(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uff00-\uffef]")
BODY_FONT = "宋体"
LATIN_FONT = "Times New Roman"


# ----------------------------------------------------------------- 基础数据
@dataclass
class Line:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    size: float
    bold: bool = False
    chars: list[tuple[float, float, str]] = field(default_factory=list)

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2

    def slice(self, x_from: float, x_to: float) -> "Line":
        """按 x 区间取这条线上属于某个单元格的字符。"""
        picked = [(x0, x1, c) for x0, x1, c in self.chars if x_from - 1 <= (x0 + x1) / 2 <= x_to + 1]
        if not picked:
            return Line(x_from, self.y0, x_from, self.y1, "", self.size, self.bold)
        text = "".join(c for _a, _b, c in picked)
        return Line(min(p[0] for p in picked), self.y0, max(p[1] for p in picked),
                    self.y1, text, self.size, self.bold, picked)


@dataclass
class Row:
    y0: float
    y1: float
    edges: list[float]
    cells: list[list[Line]] = field(default_factory=list)


@dataclass
class Table:
    x0: float
    y0: float
    x1: float
    y1: float
    rows: list[Row] = field(default_factory=list)


# ----------------------------------------------------------------- 文本整理
def tidy(text: str) -> str:
    """PDF 里中文之间常被塞进空格（字距补偿），去掉它们。"""
    if not text:
        return ""
    text = text.replace("\u00a0", " ").replace("\u3000", "　")
    text = re.sub(r"[ \t]{2,}", " ", text)
    # 仅当空格两侧都是 CJK 时才删，避免破坏英文单词
    prev = None
    while prev != text:
        prev = text
        text = CJK_RE.sub(lambda m: m.group(0), text)
        text = re.sub(r"(?<=[\u4e00-\u9fff\uff00-\uffef]) +(?=[\u4e00-\u9fff\uff00-\uffef])", "", text)
    return text.strip(" ")


def join_lines(lines: list[Line], right_edge: float | None = None) -> str:
    """把同一单元格/段落里的多行文字拼回去。

    判断依据：上一行是否明显没写满（右侧留白多）——写满了说明是自动换行，直接接上；
    留白多说明是真正的换行，用 \n 保留。
    """
    out = ""
    for i, ln in enumerate(lines):
        t = tidy(ln.text)
        if not t:
            continue
        if not out:
            out = t
            continue
        wrapped = right_edge is not None and ln_prev_right(lines[i - 1]) >= right_edge - 10
        if wrapped or (right_edge is None and _cjk_tail(out) and _cjk_head(t)):
            out += t
        else:
            out += "\n" + t
    return out


def ln_prev_right(ln: Line) -> float:
    return ln.x1


def _cjk_tail(s: str) -> bool:
    return bool(s) and bool(CJK_RE.match(s[-1]))


def _cjk_head(s: str) -> bool:
    return bool(s) and bool(CJK_RE.match(s[0]))


# ----------------------------------------------------------------- PDF 解析
def _cluster(vals: list[float], tol: float = 2.5) -> list[float]:
    vals = sorted(vals)
    groups: list[list[float]] = []
    for v in vals:
        if groups and abs(v - groups[-1][-1]) <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def page_rules(page) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
    """返回 (横线, 竖线)：横线 (y, x0, x1)，竖线 (x, y0, y1)。"""
    hs: list[tuple[float, float, float]] = []
    vs: list[tuple[float, float, float]] = []
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.y - b.y) < 1.2 and abs(a.x - b.x) >= 12:
                    hs.append(((a.y + b.y) / 2, min(a.x, b.x), max(a.x, b.x)))
                elif abs(a.x - b.x) < 1.2 and abs(a.y - b.y) >= 12:
                    vs.append(((a.x + b.x) / 2, min(a.y, b.y), max(a.y, b.y)))
            elif it[0] == "re":
                r = it[1]
                if r.height < 1.6 and r.width >= 12:
                    hs.append(((r.y0 + r.y1) / 2, r.x0, r.x1))
                elif r.width < 1.6 and r.height >= 12:
                    vs.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
    return hs, vs


def page_lines(page) -> list[Line]:
    out: list[Line] = []
    for blk in page.get_text("rawdict")["blocks"]:
        if blk.get("type") != 0:
            continue
        for ln in blk.get("lines", []):
            spans = ln.get("spans") or []
            if not spans:
                continue
            chars: list[tuple[float, float, str]] = []
            size, bold = 0.0, False
            for s in spans:
                size = max(size, float(s.get("size", 0) or 0))
                bold = bold or bool(int(s.get("flags", 0)) & 16)
                for ch in s.get("chars", []):
                    c = ch.get("c", "")
                    if not c:
                        continue
                    bx0, _by0, bx1, _by1 = ch["bbox"]
                    chars.append((float(bx0), float(bx1), c))
            if not chars:
                continue
            text = "".join(c for _a, _b, c in chars)
            if not text.strip():
                continue
            x0, y0, x1, y1 = ln["bbox"]
            out.append(Line(float(x0), float(y0), float(x1), float(y1), text,
                            size or 10.0, bold, chars))
    out.sort(key=lambda l: (round(l.y0, 1), l.x0))
    return out


def find_tables(page) -> list[Table]:
    """按网格线还原表格。行带 + 每行实际的竖线边界 → 天然支持合并单元格。"""
    hs, vs = page_rules(page)
    if not hs or not vs:
        return []

    ys = _cluster([h[0] for h in hs])
    xs = _cluster([v[0] for v in vs])

    def covering_v(x: float, y0: float, y1: float) -> bool:
        for vx, vy0, vy1 in vs:
            if abs(vx - x) <= 2.6 and vy0 <= y0 + 2.6 and vy1 >= y1 - 2.6:
                return True
        return False

    bands: list[Row] = []
    for i in range(len(ys) - 1):
        y0, y1 = ys[i], ys[i + 1]
        if not (4 <= y1 - y0 <= 320):
            continue
        edges = [x for x in xs if covering_v(x, y0, y1)]
        if len(edges) >= 2 and (edges[-1] - edges[0]) >= 40:
            bands.append(Row(y0=y0, y1=y1, edges=edges))

    tables: list[Table] = []
    for b in bands:
        if tables and abs(b.y0 - tables[-1].y1) < 3.0 and _overlap(b.edges, tables[-1].rows[-1].edges):
            t = tables[-1]
            t.rows.append(b)
            t.y1 = b.y1
            t.x0 = min(t.x0, b.edges[0])
            t.x1 = max(t.x1, b.edges[-1])
        else:
            tables.append(Table(x0=b.edges[0], y0=b.y0, x1=b.edges[-1], y1=b.y1, rows=[b]))

    # 至少 2 行才算表格，单行多半是页眉装饰线
    tables = [t for t in tables if len(t.rows) >= 2]
    return tables


def _overlap(a: list[float], b: list[float]) -> bool:
    return min(a[-1], b[-1]) - max(a[0], b[0]) > 20


def assign_text(tables: list[Table], lines: list[Line]) -> list[Line]:
    """把每一行文字放进单元格；返回没被任何表格吃掉的文字行。"""
    left: list[Line] = []
    for ln in lines:
        target = None
        for t in tables:
            if t.x0 - 3 <= ln.cx <= t.x1 + 3 and t.y0 - 3 <= ln.cy <= t.y1 + 3:
                target = t
                break
        if target is None:
            left.append(ln)
            continue
        # 表格里同一行的多个单元格常被合并成一条"线"，必须按字符 x 坐标切回各单元格
        row = None
        for r in target.rows:
            if r.y0 - 3 <= ln.cy <= r.y1 + 3:
                row = r
                break
        if row is None:
            left.append(ln)
            continue
        remaining = ln
        for i in range(len(row.edges) - 1):
            piece = ln.slice(row.edges[i], row.edges[i + 1])
            if not piece.text.strip():
                continue
            while len(row.cells) <= i:
                row.cells.append([])
            row.cells[i].append(piece)
        # 落在表格左右边界之外的字符仍然属于正文
        outer = [(a, b, c) for a, b, c in ln.chars
                 if (a + b) / 2 < row.edges[0] - 1 or (a + b) / 2 > row.edges[-1] + 1]
        if outer:
            left.append(Line(min(o[0] for o in outer), ln.y0, max(o[1] for o in outer),
                             ln.y1, "".join(o[2] for o in outer), ln.size, ln.bold, outer))
    return left


# ----------------------------------------------------------------- DOCX 渲染
def _set_cell_text(cell, text: str, size: float, bold: bool, align=None) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(1.5)
    pf.space_after = Pt(1.5)
    pf.line_spacing = 1.05
    if align is not None:
        p.alignment = align
    chunks = text.split("\n") if text else [""]
    for i, chunk in enumerate(chunks):
        if i:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.05
            if align is not None:
                p.alignment = align
        run = p.add_run(chunk)
        _style_run(run, size, bold)


def _style_run(run, size: float, bold: bool = False) -> None:
    run.font.size = Pt(max(7.5, min(size, 26)))
    run.font.bold = bold
    run.font.name = LATIN_FONT
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:ascii"), LATIN_FONT)
    rf.set(qn("w:hAnsi"), LATIN_FONT)
    rf.set(qn("w:eastAsia"), BODY_FONT)


def _cell_margins(table, l: int = 45, r: int = 45, t: int = 20, b: int = 20) -> None:
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("w:top", t), ("w:left", l), ("w:bottom", b), ("w:right", r)):
        el = OxmlElement(tag)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _table_borders(table) -> None:
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for tag in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def _fixed_layout(table) -> None:
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def _repeat_header(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    trPr.append(el)


def compose_cell_text(lines: list[Line], right_edge: float) -> str:
    """单元格内文字拼装：先把同一基线上的碎片按 x 排好拼成一行，再按换行规则接段落。"""
    if not lines:
        return ""
    rows: list[list[Line]] = []
    for ln in sorted(lines, key=lambda l: (l.cy, l.x0)):
        if rows and abs(ln.cy - rows[-1][0].cy) <= max(2.0, ln.size * 0.35):
            rows[-1].append(ln)
        else:
            rows.append([ln])
    visual: list[Line] = []
    for r in rows:
        r.sort(key=lambda l: l.x0)
        text = ""
        x0 = r[0].x0
        x1 = r[-1].x1
        for i, ln in enumerate(r):
            if i and not (_cjk_tail(text) and _cjk_head(ln.text)):
                text += " "
            text += ln.text
        visual.append(Line(x0, r[0].y0, x1, r[-1].y1, text, max(l.size for l in r),
                           any(l.bold for l in r)))
    return join_lines(visual, right_edge=right_edge)


def render_table(doc: Document, t: Table, usable_pt: float) -> int:
    """把解析出的表格写成真正的 Word 表格。"""
    all_edges = sorted({round(e, 1) for r in t.rows for e in r.edges})
    # 合并掉过近的边界
    merged: list[float] = []
    for e in all_edges:
        if not merged or e - merged[-1] > 3:
            merged.append(e)
    edges = merged
    n_cols = max(1, len(edges) - 1)

    def edge_index(x: float) -> int:
        """把一条边界线映射到 edges 的下标（最近的边界）。"""
        best, idx = 1e9, 0
        for i, e in enumerate(edges):
            d = abs(e - x)
            if d < best:
                best, idx = d, i
        return idx

    def col_of(x: float) -> int:
        return max(0, min(edge_index(x), n_cols - 1))

    width_pt = edges[-1] - edges[0]
    scale = 1.0 if width_pt <= usable_pt else usable_pt / width_pt

    table = doc.add_table(rows=len(t.rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    _fixed_layout(table)
    _table_borders(table)
    _cell_margins(table)

    # 列宽（twips）
    for i in range(n_cols):
        w = Twips(int(round((edges[i + 1] - edges[i]) * scale * 20)))
        table.columns[i].width = w
        for cell in table.columns[i].cells:
            cell.width = w

    # 逐行填内容并按行合并
    sizes: list[float] = []
    for ri, row in enumerate(t.rows):
        tr = table.rows[ri]
        blocks: list[tuple[int, int, str, list[Line]]] = []
        for ci in range(len(row.edges) - 1):
            c0 = col_of(row.edges[ci])
            # 右边界是"结束位置"，对应的列是它前一列
            c1 = max(c0, min(edge_index(row.edges[ci + 1]) - 1, n_cols - 1))
            lines = row.cells[ci] if ci < len(row.cells) else []
            if c1 > c0:
                cell = tr.cells[c0]
                for k in range(c0 + 1, c1 + 1):
                    cell = cell.merge(tr.cells[k])
            else:
                cell = tr.cells[c0]
            text = compose_cell_text(lines, row.edges[ci + 1]) if lines else ""
            if lines:
                sizes.append(max(l.size for l in lines))
            blocks.append((c0, c1, text, lines))
            head = ri == 0 and len(t.rows) > 1
            _set_cell_text(cell, text, sizes[-1] if lines else 10.5,
                           bold=head or any(l.bold for l in lines),
                           align=None)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if ri == 0 and len(t.rows) > 1:
            _repeat_header(tr)
    return len(t.rows)


NUM_RE = re.compile(r"^(?:[一二三四五六七八九十]+[、．.]|（[一二三四五六七八九十]+）|\(\d+\)|（\d+）|\d+[、．.)）])")


def render_page_text(doc: Document, lines: list[Line], page_w: float, left: float,
                     right: float, content_w: float) -> int:
    """把表格之外的文字还原成段落（按行距断段，居中/左对齐自动判断）。"""
    if not lines:
        return 0
    lines = sorted(lines, key=lambda l: (round(l.y0, 1), l.x0))
    groups: list[list[Line]] = []
    for ln in lines:
        if not groups:
            groups.append([ln])
            continue
        prev = groups[-1][-1]
        gap = ln.y0 - prev.y1
        new_para = gap > max(4.5, prev.size * 0.5)
        if not new_para and NUM_RE.match(tidy(ln.text)) and ln.x0 - left < prev.size * 1.5:
            new_para = True
        if new_para:
            groups.append([ln])
        else:
            groups[-1].append(ln)

    center = (left + right) / 2
    n_paras = 0
    for g in groups:
        joined = join_lines(g, right_edge=right)
        if not joined.strip():
            continue
        size = max(l.size for l in g)
        centers = [l.cx for l in g]
        width = max(l.x1 for l in g) - min(l.x0 for l in g)
        centered = width < content_w * 0.8 and all(abs(c - center) < max(18, size) for c in centers)
        indent = max(0.0, min(l.x0 for l in g) - left)
        for piece in joined.split("\n"):
            piece = piece.strip()
            if not piece:
                continue
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.12
            if centered:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if indent > size * 0.8:
                    pf.left_indent = Pt(indent)
            run = p.add_run(piece)
            _style_run(run, size, bold=any(l.bold for l in g))
            n_paras += 1
    return n_paras


# ----------------------------------------------------------------- 主流程
def convert(src: Path | str, out_docx: Path | str) -> tuple[bool, str, dict]:
    """PDF → DOCX（结构重建）。返回 (是否成功, 说明, 统计)。"""
    try:
        try:
            import pymupdf as fitz  # type: ignore
        except ImportError:
            import fitz  # type: ignore
    except ImportError:
        return False, "未安装 PyMuPDF", {}

    src, out_docx = Path(src), Path(out_docx)
    try:
        pdf = fitz.open(str(src))
    except Exception as e:
        return False, f"PDF 打不开：{type(e).__name__}: {e}", {}

    if len(pdf) == 0:
        return False, "PDF 没有页面", {}

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(11906), Twips(16838)  # A4
    margin = Twips(1417)  # 2.5cm
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = margin
    sec.header_distance = Twips(720)
    sec.footer_distance = Twips(560)
    usable_pt = (11906 - 1417 * 2) / 20.0  # 16cm → 453.5pt

    stats = {"pages": len(pdf), "tables": 0, "paragraphs": 0, "chars": 0, "img_only": 0}
    header_text = ""
    footer_text = ""
    page_texts: list[tuple[list[Line], list[Table]]] = []

    for page in pdf:
        lines = page_lines(page)
        tables = find_tables(page)
        rest = assign_text(tables, lines)
        page_texts.append((rest, tables))
        if not lines:
            stats["img_only"] += 1

    # 页眉页脚：贴在页上下 70pt 内、且多页重复出现的那行
    head_zone, foot_zone = 70.0, 70.0
    head_count: dict[str, int] = {}
    foot_count: dict[str, int] = {}
    for page, (rest, _t) in zip(pdf, page_texts):
        h = page.rect.height
        for ln in rest:
            key = tidy(ln.text)
            if not key:
                continue
            if ln.y1 < head_zone:
                head_count[key] = head_count.get(key, 0) + 1
            elif ln.y0 > h - foot_zone and not re.fullmatch(r"[-—\s]*\d+[-—\s]*", key):
                foot_count[key] = foot_count.get(key, 0) + 1

    need = max(2, len(pdf) // 2)
    if head_count:
        key, cnt = max(head_count.items(), key=lambda kv: kv[1])
        if cnt >= need:
            header_text = key
    if foot_count:
        key, cnt = max(foot_count.items(), key=lambda kv: kv[1])
        if cnt >= need:
            footer_text = key

    for idx, ((rest, tables), page) in enumerate(zip(page_texts, pdf)):
        body = []
        for ln in rest:
            if ln.y1 < head_zone or ln.y0 > page.rect.height - foot_zone:
                continue
            if header_text and tidy(ln.text) == header_text:
                continue
            if footer_text and tidy(ln.text) == footer_text:
                continue
            body.append(ln)
        # 按纵向顺序把段落与表格交错插入
        items: list[tuple[float, str, object]] = [(ln.y0, "text", ln) for ln in body]
        items += [(t.y0, "table", t) for t in tables]
        items.sort(key=lambda it: it[0])

        buf: list[Line] = []
        for _y, kind, obj in items:
            if kind == "text":
                buf.append(obj)  # type: ignore[arg-type]
            else:
                stats["paragraphs"] += render_page_text(doc, buf, page.rect.width, 70.9, 524.4, usable_pt)
                buf = []
                stats["tables"] += 1
                render_table(doc, obj, usable_pt)  # type: ignore[arg-type]
                stats["paragraphs"] += 1
                doc.add_paragraph().paragraph_format.space_after = Pt(2)
        stats["paragraphs"] += render_page_text(doc, buf, page.rect.width, 70.9, 524.4, usable_pt)

        # 原 PDF 的分页是有意义的（一页一个科目），照原样保留
        if idx < len(pdf) - 1:
            doc.add_page_break()

    if header_text:
        hp = sec.header.paragraphs[0]
        hp.text = ""
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _style_run(hp.add_run(header_text), 9)
    if footer_text:
        fp = sec.footer.paragraphs[0]
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _style_run(fp.add_run(footer_text), 9)

    # 覆盖度自检：重建出来的字不能明显少于 PDF 原有的字
    pdf_chars = sum(len(page.get_text("text").strip()) for page in pdf)
    doc_chars = sum(len(p.text) for p in doc.paragraphs)
    doc_chars += sum(len(c.text) for t in doc.tables for r in t.rows for c in r.cells)
    stats["chars"] = doc_chars
    stats["pdf_chars"] = pdf_chars

    if stats["img_only"] == len(pdf):
        return False, "这份 PDF 没有文本层（扫描件），需要先 OCR", stats
    if pdf_chars and doc_chars < pdf_chars * 0.55:
        return False, f"重建只覆盖了 {doc_chars / pdf_chars:.0%} 的文字，放弃该路径", stats

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_docx))
    note = (
        f"结构重建（{stats['pages']} 页 → {stats['tables']} 个真表格、"
        f"{stats['paragraphs']} 个段落，覆盖 {doc_chars}/{pdf_chars} 字）"
    )
    return True, note, stats
