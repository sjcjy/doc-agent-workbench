"""PDF → DOCX 之后的排版修复（专治 pdf2docx 产物）。

pdf2docx 是按 PDF 的绝对坐标还原版面的，所以文字位置被写成了大量"硬定位"，
一旦文档在别的环境里重排就会崩。实测一份教务文档暴露出三类问题：

1. **单元格内的绝对缩进**：例如「参考书目」这格右缩进 2020 twips、格宽 4026，
   刚好放得下；格宽一变，可用宽度就成了负数，文字逐字竖排。
2. **正文段落里的 tab 制表位与右缩进**：pdf2docx 会把跨列的一行写成"带 tab 的
   普通段落"，制表位落在版心之外，整行冲出右边界。
3. **固定行高 / 固定行距**：`trHeight hRule="exact"` 把内容裁掉半行，
   `lineRule="exact"` 让中文上下行贴死。

修复原则只有一条：**去掉绝对定位，把版面交回 Word 自己重排**。
- 单元格内：清掉 left/right/firstLine/hanging 缩进与制表位，行距改多倍；
- 正文段落：清掉右缩进与越界制表位（保留左缩进，它通常表示真实层级）；
- 表格行高：exact → atLeast（原值作为最小高度）；
- 表格超宽：仅在真的超出页宽时等比缩回，其余情况**一律不动列宽**。

不改文字、不改图片、不搬动单元格内容。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

BODY_LINE = 276          # 1.15 倍行距（240 = 单倍）
CELL_LINE = 264
EMU_PER_TWIP = 635

_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]

_TCPR_ORDER = [
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
    "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers",
    "cellIns", "cellDel", "cellMerge", "tcPrChange",
]


# ---------------------------------------------------------------- 小工具
def _ensure(parent, tag: str, order: list[str]):
    """按 OOXML 规定顺序取/建子元素（顺序错了 Word 可能判文件损坏）。"""
    el = parent.find(qn(f"w:{tag}"))
    if el is not None:
        return el
    el = OxmlElement(f"w:{tag}")
    if tag not in order:
        parent.append(el)
        return el
    idx = order.index(tag)
    for child in list(parent):
        ctag = child.tag.split("}")[-1]
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def _int(val, default: int = 0) -> int:
    try:
        return int(val)
    except (TypeError, ValueError):
        return default


def _page_width_twips(section) -> int:
    usable = int(section.page_width or 0) - int(section.left_margin or 0) - int(section.right_margin or 0)
    return max(int(usable / EMU_PER_TWIP), 4000)


def _normalize_margins(section) -> tuple[int, int]:
    """把不合理的页边距改回正常值，返回 (原可用宽, 新可用宽) 的 twips。

    Word 自带的 PDF 转换为了复原文字位置，偶尔会把左右边距设成很离谱的值
    （实测一份 21cm 的 A4，右边距被设成 7.2cm），可用宽度只剩 11.7cm，
    表格于是被挤成一团。这里只在明显异常时才动，正常文档不受影响。
    """
    old_avail = _page_width_twips(section)
    page_w = int(section.page_width or 0)
    if not page_w:
        return old_avail, old_avail
    left = int(section.left_margin or 0)
    right = int(section.right_margin or 0)
    if left + right > page_w * 0.35:          # 边距吃掉超过 1/3 版面
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
    return old_avail, _page_width_twips(section)


def _scale_table(tbl, factor: float) -> None:
    """按比例缩放整张表（列宽 + tblGrid + tblW）。"""
    rows = tbl.findall(qn("w:tr"))
    for tr in rows:
        for tc in tr.findall(qn("w:tc")):
            w = _tcw(tc)
            if w:
                _set_tcw(tc, int(w * factor))
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for c in grid.findall(qn("w:gridCol")):
            w = _int(c.get(qn("w:w")) or 0)
            if w:
                c.set(qn("w:w"), str(int(w * factor)))
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is not None:
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is not None and tblW.get(qn("w:type")) == "dxa":
            tblW.set(qn("w:w"), str(int(_int(tblW.get(qn("w:w"))) * factor)))


def _in_table(el) -> bool:
    node = el.getparent()
    while node is not None:
        tag = node.tag.split("}")[-1]
        if tag == "tc":
            return True
        if tag == "body":
            return False
        node = node.getparent()
    return False


# ---------------------------------------------------------------- 段落
def _clear_indent(pPr, *, in_cell: bool) -> bool:
    """清掉"绝对定位"式的缩进。"""
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return False
    changed = False
    if in_cell:
        # 格子内任何缩进都是 PDF 坐标残留：格子宽度才是唯一的排版依据
        for attr in ("left", "right", "firstLine", "hanging", "start", "end"):
            if ind.get(qn(f"w:{attr}")) not in (None, "0"):
                ind.set(qn(f"w:{attr}"), "0")
                changed = True
    else:
        # 正文：右缩进基本是坐标残留；左缩进常表示真实层级，保留但别越界
        for attr in ("right", "end"):
            if _int(ind.get(qn(f"w:{attr}")) or 0) > 0:
                ind.set(qn(f"w:{attr}"), "0")
                changed = True
        for attr in ("left", "start", "firstLine", "hanging"):
            if _int(ind.get(qn(f"w:{attr}")) or 0) < 0:
                ind.set(qn(f"w:{attr}"), "0")
                changed = True
    return changed


def _strip_tabs(pPr, *, in_cell: bool, avail: int) -> bool:
    """删掉越界的制表位（单元格内全删），避免整行冲出右边界。"""
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        return False
    if in_cell:
        pPr.remove(tabs)
        return True
    keep = []
    for tab in tabs.findall(qn("w:tab")):
        pos = _int(tab.get(qn("w:pos")) or 0)
        if 0 < pos <= avail:
            keep.append(tab)
    if not keep:
        pPr.remove(tabs)
        return True
    return False


def _fix_paragraph(p, *, in_cell: bool, avail: int) -> bool:
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)

    changed = False

    # 行距：exact/atLeast 的固定值会把中文压扁或裁字 → 改多倍行距
    sp = pPr.find(qn("w:spacing"))
    if sp is not None:
        rule = sp.get(qn("w:lineRule"))
        line = _int(sp.get(qn("w:line")) or 0)
        if rule in ("exact", "atLeast") or (line and line < 240):
            sp.set(qn("w:line"), str(CELL_LINE if in_cell else BODY_LINE))
            sp.set(qn("w:lineRule"), "auto")
            changed = True
        if in_cell:
            for attr in ("before", "after", "beforeLines", "afterLines"):
                if sp.get(qn(f"w:{attr}")) not in (None, "0"):
                    sp.set(qn(f"w:{attr}"), "0")
                    changed = True
        else:
            for attr, cap in (("before", "160"), ("after", "160")):
                if _int(sp.get(qn(f"w:{attr}")) or 0) > 240:
                    sp.set(qn(f"w:{attr}"), cap)
                    changed = True

    changed |= _clear_indent(pPr, in_cell=in_cell)
    changed |= _strip_tabs(pPr, in_cell=in_cell, avail=avail)

    # 文档网格会把行高吸到格子高度，导致行距莫名其妙
    snap = pPr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = _ensure(pPr, "snapToGrid", _PPR_ORDER)
        snap.set(qn("w:val"), "0")
        changed = True
    return changed


# ---------------------------------------------------------------- 表格
def _tcw(tc) -> int:
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return 0
    el = tcPr.find(qn("w:tcW"))
    return _int(el.get(qn("w:w")) or 0) if el is not None else 0


def _set_tcw(tc, width: int) -> None:
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    el = _ensure(tcPr, "tcW", _TCPR_ORDER)
    el.set(qn("w:w"), str(max(width, 0)))
    el.set(qn("w:type"), "dxa")


def _set_margin(mar, side: str, value: int) -> None:
    """按 tcMar 的固定顺序写入单边边距（顺序错了 Word 可能判文件损坏）。"""
    order = ["top", "start", "left", "bottom", "end", "right",
             "insideH", "insideV", "tl2br", "tr2bl"]
    el = mar.find(qn(f"w:{side}"))
    if el is None:
        el = OxmlElement(f"w:{side}")
        idx = order.index(side) if side in order else len(order)
        placed = False
        for child in list(mar):
            ctag = child.tag.split("}")[-1]
            if ctag in order and order.index(ctag) > idx:
                child.addprevious(el)
                placed = True
                break
        if not placed:
            mar.append(el)
    el.set(qn("w:w"), str(value))
    el.set(qn("w:type"), "dxa")


def _fix_table(tbl, avail: int, stats: dict, old_avail: int = 0) -> None:
    rows = tbl.findall(qn("w:tr"))
    if not rows:
        return

    # 不做跨行合并：pdf2docx 的 vMerge 起点本身就错行，强行合并会把内容搬进
    # 相邻逻辑行、并把行高撑破（实测文字会压到边框上），所以只做"去定位残留"。

    # 1) 行高：固定值 → 最小高度（内容多时自动撑开，不再裁字）
    for tr in rows:
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            continue
        th = trPr.find(qn("w:trHeight"))
        if th is not None and th.get(qn("w:hRule")) != "atLeast":
            th.set(qn("w:hRule"), "atLeast")
            stats["rows"] += 1
        cs = trPr.find(qn("w:cantSplit"))
        if cs is not None:
            trPr.remove(cs)

    # 2) 单元格内部：清缩进/制表位，允许换行
    for tr in rows:
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is not None:
                no_wrap = tcPr.find(qn("w:noWrap"))
                if no_wrap is not None:
                    tcPr.remove(no_wrap)
                # pdf2docx / Word 转换常把单元格边距设成 0，文字直接贴着边框线
                mar = _ensure(tcPr, "tcMar", _TCPR_ORDER)
                _set_margin(mar, "top", 30)
                _set_margin(mar, "bottom", 30)
                _set_margin(mar, "left", 60)
                _set_margin(mar, "right", 60)
            for p in tc.findall(qn("w:p")):
                if _fix_paragraph(p, in_cell=True, avail=avail):
                    stats["cell_paras"] += 1

    # 3) 列宽同步：
    #    a. 表格宽度正好等于"旧可用宽"→ 说明它是按异常页边距铺满的，随页边距一起放大；
    #    b. 表格比页面还宽（含表格缩进）→ 等比缩回页宽内；
    #    其余情况一律不动（正常表格不该被重排）。
    widest = max(sum(_tcw(tc) for tc in tr.findall(qn("w:tc"))) for tr in rows)
    tblPr = tbl.find(qn("w:tblPr"))
    ind_el = tblPr.find(qn("w:tblInd")) if tblPr is not None else None
    indent = 0
    if ind_el is not None and ind_el.get(qn("w:type")) in (None, "dxa"):
        indent = max(0, _int(ind_el.get(qn("w:w")) or 0))

    grow = bool(old_avail and old_avail != avail and widest
                and abs(widest - old_avail) <= old_avail * 0.04)
    overflow = indent + widest > int(avail * 1.02)
    if grow or overflow:
        # 表格缩进会把整表推出右边距（Word 转换常见），先归零再定宽
        if indent and ind_el is not None:
            ind_el.set(qn("w:w"), "0")
        if widest:
            _scale_table(tbl, avail / widest)
            stats["scaled"] += 1


# ---------------------------------------------------------------- 入口
def repair_document(doc) -> dict:
    """就地修复一个已打开的 python-docx Document，返回统计信息。"""
    stats = {"tables": 0, "rows": 0, "cell_paras": 0, "paragraphs": 0,
             "scaled": 0, "page_width_twips": 0}
    try:
        old_avail, avail = _normalize_margins(doc.sections[0])
    except (IndexError, AttributeError):
        old_avail, avail = 0, 9000
    stats["page_width_twips"] = avail

    body = doc.element.body
    for tbl in list(body.iter(qn("w:tbl"))):
        _fix_table(tbl, avail, stats, old_avail)
        stats["tables"] += 1

    for p in body.iter(qn("w:p")):
        if _in_table(p):
            continue  # 已在 _fix_table 里按单元格处理
        if _fix_paragraph(p, in_cell=False, avail=avail):
            stats["paragraphs"] += 1
    return stats


def repair_docx_file(path: str | Path) -> dict:
    """打开文件 → 修复 → 存回原路径。"""
    path = Path(path)
    doc = Document(str(path))
    stats = repair_document(doc)
    doc.save(str(path))
    return stats
