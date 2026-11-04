from __future__ import annotations

"""python-docx 底层操作工具箱。

所有工具共用这里的原子操作，避免每个工具各写一套导致行为不一致。
两个关键坑在这里统一解决：
1. 跨 run 文本：中文文档常把一段话拆成多个 run，直接按 run 查找会漏。
   解决方案是把整段文本合并后再匹配替换（replace_in_paragraph）。
2. 中文字体：python-docx 的 font.name 只写 ascii/hAnsi，中文实际走 eastAsia，
   不单独设置的话 Word 里看到的字体不会变（set_run_font）。
"""

import copy
from typing import Iterable, Iterator

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

EAST_ASIA = "eastAsia"


# ---------------------------------------------------------------- 打开与遍历
def open_docx(path: str) -> Document:
    """只读打开 DOCX。调用方负责保证路径已过 safety 校验。"""
    return Document(str(path))


def iter_block_paragraphs(doc: Document) -> Iterator[Paragraph]:
    """按文档体顺序遍历段落（含表格内的段落排在表格位置）。"""
    yield from doc.paragraphs


def heading_level(p: Paragraph) -> int:
    """返回标题级别：0 表示非标题，1-9 表示 Heading 1-9 / 标题 1-9。"""
    name = (p.style.name if p.style is not None else "") or ""
    low = name.lower().replace(" ", "")
    for i in range(1, 10):
        if low in (f"heading{i}", f"标题{i}", f"head{i}"):
            return i
    # 兼容「Heading 1 Char」之类
    for i in range(1, 10):
        if low.startswith(f"heading{i}") or low.startswith(f"标题{i}"):
            return i
    return 0


def is_heading(p: Paragraph) -> bool:
    return heading_level(p) > 0


def para_text(p: Paragraph) -> str:
    return "".join(r.text for r in p.runs) if p.runs else p.text


def non_empty_paragraphs(doc: Document) -> list[tuple[int, Paragraph]]:
    return [(i, p) for i, p in enumerate(doc.paragraphs) if para_text(p).strip()]


def word_count(doc: Document) -> int:
    """粗略字数统计：中文按字，英文按词。"""
    import re as _re
    total = 0
    for p in doc.paragraphs:
        t = para_text(p)
        cn = len(_re.findall(r"[\u4e00-\u9fa5]", t))
        en = len(_re.findall(r"[A-Za-z0-9]+", t))
        total += cn + en
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    tt = para_text(p)
                    total += len(_re.findall(r"[\u4e00-\u9fa5]", tt)) + len(_re.findall(r"[A-Za-z0-9]+", tt))
    return total


# ---------------------------------------------------------------- 文本替换
def merge_runs(p: Paragraph) -> None:
    """把段落内的多个 run 合并成一个，便于可靠匹配与替换。

    保留第一个 run 的原有格式，清掉其余 run。
    """
    runs = p.runs
    if len(runs) <= 1:
        return
    full = "".join(r.text for r in runs)
    first = runs[0]
    first.text = full
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def replace_in_paragraph(p: Paragraph, old: str, new: str, occurrence: int = 1) -> int:
    """在单个段落里替换文本，返回实际替换次数。

    occurrence: 1 表示只换第一处，0 表示全部替换，n>1 表示换第 n 处。
    采用「先合并 run 再替换」，规避跨 run 匹配不到的经典问题。
    """
    if not old:
        return 0
    merge_runs(p)
    if not p.runs:
        return 0
    text = p.runs[0].text
    if old not in text:
        return 0

    if occurrence == 0:
        p.runs[0].text = text.replace(old, new)
        return text.count(old)
    idxs = [i for i in range(len(text)) if text.startswith(old, i)]
    k = min(occurrence, len(idxs))
    start = idxs[k - 1]
    p.runs[0].text = text[:start] + new + text[start + len(old):]
    return 1


def replace_in_table(table: Table, old: str, new: str, occurrence: int = 1) -> int:
    """在表格所有单元格里替换文本。"""
    n = 0
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                n += replace_in_paragraph(p, old, new, 0 if occurrence == 0 else occurrence)
    return n


# ---------------------------------------------------------------- 字体与段落格式
def _ensure_rpr(run) -> None:
    run.font.name = run.font.name or None  # 触发 rPr 创建
    rpr = run._element.get_or_add_rPr()
    return rpr


def set_run_font(
    run,
    cn_font: str | None = None,
    en_font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color_hex: str | None = None,
) -> None:
    """设置 run 的字体。cn_font 会写入 eastAsia，保证 Word 中中文真的变化。"""
    if cn_font:
        run.font.name = cn_font
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), cn_font)
        rfonts.set(qn("w:hAnsi"), cn_font)
        rfonts.set(qn("w:eastAsia"), cn_font)
    if en_font:
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), en_font)
        rfonts.set(qn("w:hAnsi"), en_font)
    if size_pt:
        run.font.size = Pt(size_pt)
        # 中文还需要 szCs，否则 Word 里部分场景不跟随
        rpr = run._element.get_or_add_rPr()
        sz = rpr.find(qn("w:szCs"))
        if sz is None:
            sz = rpr.makeelement(qn("w:szCs"), {})
            rpr.append(sz)
        sz.set(qn("w:val"), str(int(size_pt * 2)))
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color_hex:
        try:
            run.font.color.rgb = RGBColor.from_string(color_hex.lstrip("#"))
        except ValueError:
            pass


def set_para_font(p: Paragraph, cn_font=None, en_font=None, size_pt=None, bold=None) -> None:
    """设置整段字体。空段落没有 run，此时补一个空 run 承载格式。"""
    if not p.runs:
        p.add_run("")
    for r in p.runs:
        set_run_font(r, cn_font=cn_font, en_font=en_font, size_pt=size_pt, bold=bold)


def align_paragraph(p: Paragraph, alignment: str) -> bool:
    """设置对齐：left/center/right/justify。返回是否真的改变了。"""
    mapping = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    target = mapping.get((alignment or "").lower())
    if target is None:
        return False
    if p.paragraph_format.alignment == target:
        return False
    p.paragraph_format.alignment = target
    return True


def set_line_spacing(p: Paragraph, spacing: float, exact: bool = False) -> bool:
    """设置行距。spacing=1.5 表示 1.5 倍；exact=True 时按磅值固定。"""
    pf = p.paragraph_format
    if exact:
        from docx.shared import Pt as _Pt
        new_val = _Pt(spacing)
        if pf.line_spacing == new_val and pf.line_spacing_rule:
            return False
        pf.line_spacing = new_val
        from docx.enum.text import WD_LINE_SPACING
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    else:
        if pf.line_spacing == spacing:
            return False
        pf.line_spacing = spacing
    return True


def set_first_line_indent(p: Paragraph, chars: float | None = None, pt: float | None = None) -> bool:
    """设置首行缩进。chars 按「字符数」换算（1 字符 = 该段字号），pt 直接按磅。

    python-docx 的 first_line_indent 正值即为 Word 的「首行缩进」。
    """
    pf = p.paragraph_format
    if pt is not None:
        value = Pt(pt)
    elif chars is not None:
        size = None
        for r in p.runs:
            if r.font.size:
                size = r.font.size.pt
                break
        base = size or 12.0
        value = Pt(base * float(chars))
    else:
        return False
    if pf.first_line_indent == value:
        return False
    pf.first_line_indent = value
    return True


def set_space_after(p: Paragraph, pt: float | None) -> bool:
    if pt is None:
        return False
    v = Pt(pt)
    if p.paragraph_format.space_after == v:
        return False
    p.paragraph_format.space_after = v
    return True


# ---------------------------------------------------------------- 段落增删
def _new_paragraph_before(doc: Document, para_el, text: str = "", style=None) -> Paragraph:
    el = copy.deepcopy(para_el)
    para_el.addprevious(el)
    return Paragraph(el, doc)


def insert_paragraph_after(anchor: Paragraph, text: str = "", style: str | None = None) -> Paragraph:
    """在指定段落后插入新段落，返回新段落对象。"""
    doc = anchor._parent
    new_el = copy.deepcopy(anchor._element)
    anchor._element.addnext(new_el)
    p = Paragraph(new_el, doc)
    # 清空文本但继承样式
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    if style:
        try:
            p.style = style
        except KeyError:
            pass
    if text:
        p.add_run(text)
    else:
        p.add_run("")
    return p


def append_paragraph(doc: Document, text: str = "", style: str | None = None) -> Paragraph:
    p = doc.add_paragraph()
    if style and style in _available_styles(doc):
        p.style = style
    if text:
        p.add_run(text)
    return p


def _available_styles(doc: Document) -> set[str]:
    try:
        return {s.name for s in doc.styles}
    except Exception:
        return set()


def resolve_style(doc: Document, want: str, fallbacks: Iterable[str] = ()) -> str | None:
    """按名称找样式，找不到就用回退列表里的第一个可用样式。"""
    names = _available_styles(doc)
    for cand in (want, *fallbacks):
        if cand in names:
            return cand
    return None


def add_heading(doc: Document, text: str, level: int = 1) -> Paragraph:
    """添加标题，优先用 Heading N 样式，缺失时退化为加粗段落。"""
    style = None
    for cand in (f"Heading {level}", f"标题 {level}", f"Heading {level} Char", "Title"):
        style = resolve_style(doc, cand)
        if style:
            break
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    if level == 1:
        set_run_font(r, size_pt=18, bold=True)
    elif level == 2:
        set_run_font(r, size_pt=15, bold=True)
    else:
        set_run_font(r, size_pt=13, bold=True)
    return p


# ---------------------------------------------------------------- 读取辅助
def document_outline(doc: Document, max_items: int = 60) -> list[dict]:
    """提取文档大纲：标题级别 + 文本 + 段序号。"""
    items = []
    for i, p in enumerate(doc.paragraphs):
        lv = heading_level(p)
        t = para_text(p).strip()
        if not t:
            continue
        if lv > 0 or (i < 3 and t):
            kind = f"H{lv}" if lv else ("Title" if i == 0 else "Body")
            items.append({"index": i, "level": lv, "kind": kind, "text": t[:80]})
        if len(items) >= max_items:
            break
    return items


def preview_text(doc: Document, limit: int = 600) -> str:
    buf = []
    n = 0
    for _, p in non_empty_paragraphs(doc):
        buf.append(para_text(p).strip())
        n += len(buf[-1])
        if n >= limit:
            break
    return "\n".join(buf)[:limit]
