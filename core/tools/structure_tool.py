from __future__ import annotations

"""工具六：文档结构操作（structure）。

四个动作：
- merge           把多份 DOCX 按顺序合并成一份（可选分页、可选加来源标注）
- split           按标题层级拆成多份（也可只从某个标题处切开）
- extract_tables  把文档里的表格导出成 Excel（一个表一个工作表）
- diff            对比两版文档，输出差异报告（新增 / 删除 / 修改）

铁律与其它工具一致：永远读源文件、写新文件，源文件路径不作为保存目标。

几个实现上的坑，在这里统一处理：
1. 合并不能只搬文字：段落里可能带图片、超链接，它们的 rId 属于源文档的关系表，
   直接 deepcopy 到目标文档会变成断链，Word 打开时报「无法读取的内容」。
   所以复制元素后要重建关系（_rebase_rels）。
2. 各文档的样式 ID 不一定都在目标文档里注册过，标题会退化成正文。
   合并前先把源文档的样式定义补进目标文档（_merge_styles）。
3. doc.paragraphs 的顺序和 body 元素顺序不一致（表格会插在中间）。
   拆分必须按 body 子元素的真实顺序切片，否则会切错位置。
"""

import copy
import json
import re
import time
from difflib import SequenceMatcher
from pathlib import Path

from pydantic import BaseModel, Field
from typing_extensions import Literal

from core.ctx import get_session_id
from core.logger import log_call
from core.models import (
    CODE_INTERNAL_ERROR,
    CODE_INVALID_INPUT,
    CODE_NO_CHANGE,
    CODE_NOT_FOUND_IN_DOC,
    CODE_UNSUPPORTED_FORMAT,
    ToolResult,
)
from core.safety import InputGuard, build_output_path, validate_input_file
from core.tools import _docxkit as dk

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

TOOL_NAME = "structure_document"
TOOL_DESC = (
    "对 DOCX 做结构级操作，输出新文件（源文件保持只读）。"
    "action=merge 把多份文档合并成一份，other_paths 填其余文件路径（用 | 分隔）；"
    "action=split 按标题拆分成多份，level 指定按几级标题拆（默认 1），"
    "也可用 anchor 指定从含该文字的标题处切开；"
    "action=extract_tables 把文档里的表格导出成 Excel，一个表一个工作表；"
    "action=diff 对比两版文档，other_paths 填新版文件路径，输出差异报告。"
)

StructureAction = Literal["merge", "split", "extract_tables", "diff"]

MAX_MERGE_FILES = 20


class StructureArgs(BaseModel):
    path: str = Field(description="主文件绝对路径（只读，绝不会被覆盖）")
    action: StructureAction = Field(description="结构操作动作")
    other_paths: str = Field(
        default="",
        description="merge 时的其余文件路径、diff 时的新版文件路径，多个用 | 或换行分隔",
    )
    level: int = Field(default=1, description="split 时按几级标题拆分，1 表示只按一级标题")
    anchor: str = Field(default="", description="split 时只从包含该文字的标题处切开")
    add_page_break: bool = Field(default=True, description="merge 时每份文档之间是否分页")
    mark_source: bool = Field(default=False, description="merge 时是否给每份内容加一行来源文件名标注")


# ---------------------------------------------------------------- 通用辅助
def _split_paths(raw: str) -> list[str]:
    """把用户给的多文件串拆成路径列表。支持 | 、换行、分号。

    不用逗号，因为 Windows 文件名里可能真的带逗号。
    """
    if not raw:
        return []
    out: list[str] = []
    for piece in re.split(r"[|\n\r;]+", str(raw)):
        s = piece.strip().strip('"').strip("'").strip()
        if s:
            out.append(s)
    return out


def _merge_styles(dst: Document, src: Document) -> int:
    """把源文档里目标文档没有的样式定义补进去，否则 Heading 会退化成正文。"""
    try:
        dst_root = dst.styles.element
        have = {s.get(qn("w:styleId")) for s in dst_root.findall(qn("w:style"))}
        added = 0
        for st in src.styles.element.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            if not sid or sid in have:
                continue
            dst_root.append(copy.deepcopy(st))
            have.add(sid)
            added += 1
        return added
    except Exception:
        return 0


def _rebase_rels(dst: Document, el, src: Document) -> int:
    """把复制过来的元素引用的图片/超链接关系搬到目标文档，并改写 rId。

    返回没能迁移的引用数量（调用方据此清理断链，保证文件能被 Word 打开）。
    """
    rid_map: dict[str, str] = {}
    lost = 0
    attrs = (qn("r:embed"), qn("r:link"), qn("r:id"))
    for node in el.iter():
        for attr in attrs:
            rid = node.get(attr)
            if not rid:
                continue
            if rid in rid_map:
                node.set(attr, rid_map[rid])
                continue
            try:
                rel = src.part.rels[rid]
            except KeyError:
                node.attrib.pop(attr, None)
                lost += 1
                continue
            try:
                if rel.is_external:
                    new_rid = dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                else:
                    new_rid = dst.part.relate_to(rel.target_part, rel.reltype)
                rid_map[rid] = new_rid
                node.set(attr, new_rid)
            except Exception:
                node.attrib.pop(attr, None)
                lost += 1
    return lost


def _strip_broken_pictures(doc: Document) -> int:
    """删掉引用了不存在关系的图片，避免 Word 报「无法读取的内容」。"""
    removed = 0
    for holder in (qn("w:drawing"), qn("w:pict")):
        for node in list(doc.element.body.iter(holder)):
            broken = False
            for blip in node.iter():
                tag = blip.tag if isinstance(blip.tag, str) else ""
                if tag.endswith("}blip") and not blip.get(qn("r:embed")) and not blip.get(qn("r:link")):
                    broken = True
                    break
            if not broken:
                continue
            run = node.getparent()
            if run is None:
                continue
            run.remove(node)
            removed += 1
    return removed


def _copy_page_setup(dst: Document, src: Document) -> None:
    """页面尺寸与页边距跟随源文档。"""
    try:
        a, b = src.sections[0], dst.sections[0]
        b.page_width = a.page_width
        b.page_height = a.page_height
        b.orientation = a.orientation
        b.left_margin = a.left_margin
        b.right_margin = a.right_margin
        b.top_margin = a.top_margin
        b.bottom_margin = a.bottom_margin
    except Exception:
        pass


def _add_page_break(doc: Document) -> None:
    from docx.enum.text import WD_BREAK

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _body_sectpr(doc: Document):
    return doc.element.body.find(qn("w:sectPr"))


def _attach(dst: Document, el, sect_pr) -> None:
    """把元素挂到 dst 的 sectPr 之前（保持它在文末、页面设置之前）。"""
    if sect_pr is not None:
        sect_pr.addprevious(el)
    else:
        dst.element.body.append(el)


def _iter_blocks(doc: Document):
    """按 body 里子元素的真实顺序遍历：('p', Paragraph) / ('tbl', element) / ('other', element)。"""
    for child in doc.element.body:
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag == qn("w:p"):
            yield "p", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "tbl", child
        else:
            yield "other", child


def _append_document(dst: Document, src: Document, add_break: bool, label: str = "") -> tuple[int, int]:
    """把 src 的正文整体搬到 dst 末尾。返回 (元素数, 丢失的引用数)。"""
    sect_pr = _body_sectpr(dst)
    if add_break:
        _add_page_break(dst)
        sect_pr = _body_sectpr(dst)

    if label:
        from docx.shared import Pt

        lp = dst.add_paragraph()
        lr = lp.add_run(label)
        dk.set_run_font(lr, size_pt=9, color_hex="808080")
        lp.paragraph_format.space_after = Pt(4)
        sect_pr = _body_sectpr(dst)

    count = 0
    lost = 0
    for child in list(src.element.body):
        if child.tag == qn("w:sectPr"):
            continue
        el = copy.deepcopy(child)
        lost += _rebase_rels(dst, el, src)
        _attach(dst, el, sect_pr)
        count += 1
    return count, lost


# ---------------------------------------------------------------- merge
def _do_merge(paths: list[Path], out_path: Path, add_page_break: bool, mark_source: bool) -> ToolResult:
    tool, action = "structure", "merge"
    t0 = time.perf_counter()

    dst = Document()
    first = dk.open_docx(paths[0])
    _copy_page_setup(dst, first)

    report = []
    total_lost = 0
    total_elems = 0
    for i, p in enumerate(paths):
        src = first if i == 0 else dk.open_docx(p)
        style_added = _merge_styles(dst, src)
        n, lost = _append_document(
            dst, src, add_break=add_page_break and i > 0,
            label=(f"——《{p.stem}》" if mark_source and i > 0 else ""),
        )
        total_elems += n
        total_lost += lost
        report.append({
            "file": p.name,
            "elements": n,
            "paragraphs": len(src.paragraphs),
            "tables": len(src.tables),
            "styles_added": style_added,
            "lost_refs": lost,
        })

    if total_lost:
        _strip_broken_pictures(dst)

    dst.save(str(out_path))
    elapsed = int((time.perf_counter() - t0) * 1000)

    msg = f"已合并 {len(paths)} 份文档，共 {total_elems} 个内容块"
    if total_lost:
        msg += f"；其中 {total_lost} 处图片/链接因跨文件引用无法迁移，已略去"

    return ToolResult.ok_result(
        tool, action, msg,
        output_file=str(out_path),
        changes=[{"type": "merge", **r} for r in report],
        data={"source": str(paths[0]), "files": [str(p) for p in paths],
              "file_count": len(paths), "lost_refs": total_lost},
        elapsed_ms=elapsed,
    )


# ---------------------------------------------------------------- split
def _do_split(path: Path, level: int, anchor: str) -> ToolResult:
    tool, action = "structure", "split"
    t0 = time.perf_counter()

    src = dk.open_docx(path)
    blocks = list(_iter_blocks(src))
    if not blocks:
        return ToolResult.fail(tool, action, CODE_NOT_FOUND_IN_DOC, f"《{path.name}》里没有任何内容")

    a = (anchor or "").strip()
    cut: list[tuple[int, str]] = []
    used_level = max(1, int(level))
    auto_note = ""

    def _collect(max_level: int) -> list[tuple[int, str]]:
        found: list[tuple[int, str]] = []
        for i, (kind, obj) in enumerate(blocks):
            if kind != "p":
                continue
            txt = dk.para_text(obj).strip()
            if not txt:
                continue
            lv = dk.heading_level(obj)
            if a:
                if a in txt and lv > 0:
                    found.append((i, txt))
            elif 0 < lv <= max_level:
                found.append((i, txt))
        return found

    cut = _collect(used_level)

    # 很多文档标题层级不规范（通篇只有「标题 2」），按死 level=1 会一份都拆不出来。
    # 此时自动回退到文档里实际存在的最小标题级别，并在结果里说明，不静默改变行为。
    if not cut and not a:
        available = sorted({
            dk.heading_level(obj) for kind, obj in blocks
            if kind == "p" and dk.heading_level(obj) > 0
        })
        if available:
            used_level = available[0]
            cut = _collect(used_level)
            if cut:
                auto_note = f"（文档没有 {max(1, int(level))} 级标题，已按 {used_level} 级标题拆分）"

    if not cut:
        hint = (
            f"没有找到包含「{a}」的标题" if a else
            f"文档里没有任何标题样式（一级到 {level} 级）"
        )
        return ToolResult.fail(
            tool, action, CODE_NOT_FOUND_IN_DOC,
            f"{hint}，无法拆分。请确认文档标题使用了「标题 1 / Heading 1」这类样式，"
            "可以先用格式工具把它们设为标题。",
            data={"paragraphs": len(src.paragraphs), "tables": len(src.tables)},
            elapsed_ms=int((time.perf_counter() - t0) * 1000),
        )

    # 切分：第一个标题之前的内容单独成篇（前言），其余每个标题起一段
    segs: list[tuple[list, str]] = []
    if cut[0][0] > 0:
        segs.append((blocks[: cut[0][0]], "前言"))
    for j, (start, title) in enumerate(cut):
        end = cut[j + 1][0] if j + 1 < len(cut) else len(blocks)
        segs.append((blocks[start:end], title))

    base = build_output_path(path, "split", ".docx")
    files: list[str] = []
    parts: list[dict] = []
    for idx, (seg_blocks, title) in enumerate(segs, 1):
        part = Document()
        _copy_page_setup(part, src)
        _merge_styles(part, src)
        sect_pr = _body_sectpr(part)
        for _kind, obj in seg_blocks:
            el = obj._element if isinstance(obj, Paragraph) else obj
            new_el = copy.deepcopy(el)
            _rebase_rels(part, new_el, src)
            _attach(part, new_el, sect_pr)
        safe = re.sub(r"[^\w\u4e00-\u9fa5]+", "_", title)[:24] or f"part{idx}"
        target = base.with_name(f"{base.stem}_{idx:02d}_{safe}{base.suffix}")
        while target.exists():
            target = target.with_name(f"{base.stem}_{idx:02d}_{safe}_{idx}{base.suffix}")
        part.save(str(target))
        files.append(str(target))
        parts.append({"index": idx, "title": title, "file": target.name,
                      "blocks": len(seg_blocks)})

    elapsed = int((time.perf_counter() - t0) * 1000)
    return ToolResult.ok_result(
        tool, action,
        f"已拆成 {len(files)} 份：{'、'.join(p['title'] for p in parts[:4])}"
        + ("…" if len(parts) > 4 else "") + auto_note,
        output_file=files[0],
        changes=[{"type": "split", **p} for p in parts],
        data={"source": str(path), "files": files, "count": len(files), "parts": parts},
        elapsed_ms=elapsed,
    )


# ---------------------------------------------------------------- extract_tables
def _do_extract_tables(path: Path, out_path: Path) -> ToolResult:
    tool, action = "structure", "extract_tables"
    t0 = time.perf_counter()

    src = dk.open_docx(path)
    tables = src.tables
    if not tables:
        return ToolResult.fail(
            tool, action, CODE_NOT_FOUND_IN_DOC,
            f"《{path.name}》里没有表格。如果是 PDF，请先转成 Word 再提取。",
            data={"tables": 0},
            elapsed_ms=int((time.perf_counter() - t0) * 1000),
        )

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    except ImportError:
        return ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT,
                               "缺少 openpyxl，无法生成 Excel。请执行 pip install openpyxl")

    wb = Workbook()
    wb.remove(wb.active)

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="EAF2FB")
    head_font = Font(bold=True, size=11)
    body_align = Alignment(vertical="center", wrap_text=True)

    summary: list[dict] = []
    for ti, t in enumerate(tables, 1):
        title = f"表{ti}"
        name_taken = set(wb.sheetnames)
        if title in name_taken:
            k = 2
            while f"{title}_{k}" in name_taken:
                k += 1
            title = f"{title}_{k}"
        ws = wb.create_sheet(title[:31])

        max_col = 0
        row_count = 0
        for ri, row in enumerate(t.rows, 1):
            seen: dict[int, int] = {}
            for ci, cell in enumerate(row.cells, 1):
                key = id(cell._tc)
                if key in seen:
                    continue
                # 计算这个单元格在视觉上横跨几列（合并单元格）
                span = 1
                for cj in range(ci + 1, len(row.cells) + 1):
                    if id(row.cells[cj - 1]._tc) == key:
                        span += 1
                    else:
                        break
                seen[key] = span
                value = cell.text.strip()
                c = ws.cell(row=ri, column=ci, value=value or None)
                c.border = border
                c.alignment = body_align
                if ri == 1:
                    c.font = head_font
                    c.fill = head_fill
                if span > 1:
                    ws.merge_cells(start_row=ri, start_column=ci,
                                   end_row=ri, end_column=ci + span - 1)
                max_col = max(max_col, ci)
            row_count = ri

        # 列宽按内容估算（中文按 2 个字符宽）
        for ci in range(1, max_col + 1):
            width = 8
            for ri in range(1, min(row_count, 60) + 1):
                v = ws.cell(row=ri, column=ci).value
                if v is None:
                    continue
                s = str(v)
                w = sum(2 if ord(ch) > 127 else 1 for ch in s[:40]) + 2
                width = max(width, min(w, 50))
            ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = width
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 22

        summary.append({"index": ti, "sheet": ws.title, "rows": row_count, "cols": max_col})

    try:
        wb.save(str(out_path))
    except PermissionError:
        return ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                               "Excel 保存失败，文件可能正在被 Excel 打开，请关闭后重试")

    elapsed = int((time.perf_counter() - t0) * 1000)
    total_rows = sum(s["rows"] for s in summary)
    return ToolResult.ok_result(
        tool, action,
        f"已导出 {len(summary)} 个表格到 Excel（共 {total_rows} 行）",
        output_file=str(out_path),
        changes=[{"type": "extract_table", **s} for s in summary],
        data={"source": str(path), "tables": summary},
        elapsed_ms=elapsed,
    )


# ---------------------------------------------------------------- diff
def _do_diff(old_path: Path, new_path: Path, out_path: Path) -> ToolResult:
    tool, action = "structure", "diff"
    t0 = time.perf_counter()

    a = dk.open_docx(old_path)
    b = dk.open_docx(new_path)
    a_texts = [dk.para_text(p).strip() for p in a.paragraphs]
    b_texts = [dk.para_text(p).strip() for p in b.paragraphs]
    a_texts = [x for x in a_texts if x]
    b_texts = [x for x in b_texts if x]

    sm = SequenceMatcher(None, a_texts, b_texts, autojunk=False)
    rows: list[dict] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if tag == "replace":
            n = max(i2 - i1, j2 - j1)
            for k in range(n):
                old = a_texts[i1 + k] if i1 + k < i2 else ""
                new = b_texts[j1 + k] if j1 + k < j2 else ""
                if old and new:
                    sim = SequenceMatcher(None, old, new, autojunk=False).ratio()
                    if sim >= 0.45:
                        rows.append({"kind": "修改", "old": old, "new": new,
                                     "similarity": round(sim, 3)})
                    else:
                        # 两段几乎不沾边，说成「修改」会误导，拆成删除 + 新增更准确
                        rows.append({"kind": "删除", "old": old, "new": ""})
                        rows.append({"kind": "新增", "old": "", "new": new})
                elif old:
                    rows.append({"kind": "删除", "old": old, "new": ""})
                else:
                    rows.append({"kind": "新增", "old": "", "new": new})
        elif tag == "delete":
            for k in range(i1, i2):
                rows.append({"kind": "删除", "old": a_texts[k], "new": ""})
        else:  # insert
            for k in range(j1, j2):
                rows.append({"kind": "新增", "old": "", "new": b_texts[k]})

    n_add = sum(1 for r in rows if r["kind"] == "新增")
    n_del = sum(1 for r in rows if r["kind"] == "删除")
    n_mod = sum(1 for r in rows if r["kind"] == "修改")
    ratio = sm.ratio()

    if not rows:
        return ToolResult.fail(
            tool, action, CODE_NO_CHANGE,
            "两份文档的段落内容完全一致，没有差异",
            data={"similarity": round(ratio, 4)},
            elapsed_ms=int((time.perf_counter() - t0) * 1000),
        )

    doc = Document()
    dk.add_heading(doc, "文档对比报告", 1)
    p = doc.add_paragraph()
    lines = [
        f"旧版：{old_path.name}",
        f"新版：{new_path.name}",
        f"对比结果：新增 {n_add} 段、删除 {n_del} 段、修改 {n_mod} 段，"
        f"段落相似度 {ratio:.1%}",
    ]
    for i, line in enumerate(lines):
        run = p.add_run(line)
        dk.set_run_font(run, size_pt=10.5)
        if i < len(lines) - 1:
            run.add_break()

    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid" if dk.resolve_style(doc, "Table Grid") else None
    hdr = ["序号", "变化", "原文", "新文"]
    for i, h in enumerate(hdr):
        cell = t.rows[0].cells[i]
        cell.text = ""
        rr = cell.paragraphs[0].add_run(h)
        dk.set_run_font(rr, size_pt=10.5, bold=True)

    from docx.shared import Pt

    for idx, row in enumerate(rows, 1):
        cells = t.add_row().cells
        vals = [str(idx), row["kind"], row["old"] or "—", row["new"] or "—"]
        for i, v in enumerate(vals):
            cells[i].text = ""
            para = cells[i].paragraphs[0]
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run(v)
            color = "C00000" if row["kind"] == "删除" else ("1F7A1F" if row["kind"] == "新增" else "000000")
            dk.set_run_font(run, size_pt=10, color_hex=color)

    doc.save(str(out_path))
    elapsed = int((time.perf_counter() - t0) * 1000)
    return ToolResult.ok_result(
        tool, action,
        f"对比完成：新增 {n_add} 段、删除 {n_del} 段、修改 {n_mod} 段（相似度 {ratio:.1%}）",
        output_file=str(out_path),
        changes=[{"type": "diff", **r} for r in rows[:200]],
        data={"source": str(old_path), "other": str(new_path),
              "added": n_add, "deleted": n_del, "modified": n_mod,
              "similarity": round(ratio, 4), "count": len(rows)},
        elapsed_ms=elapsed,
    )


# ---------------------------------------------------------------- 入口
def run(
    path: str,
    action: str,
    other_paths: str = "",
    level: int = 1,
    anchor: str = "",
    add_page_break: bool = True,
    mark_source: bool = False,
    session_id: str = "",
) -> ToolResult:
    session_id = session_id or get_session_id()
    t0 = time.perf_counter()
    tool = "structure"
    args = {"path": path, "action": action, "other_paths": other_paths,
            "level": level, "anchor": anchor}

    if action not in ("merge", "split", "extract_tables", "diff"):
        r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                            f"不支持的结构操作：{action}（可选 merge/split/extract_tables/diff）",
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
        return r

    ok_path, err = validate_input_file(path, tool, action)
    if err is not None:
        err.elapsed_ms = int((time.perf_counter() - t0) * 1000)
        log_call(session_id, "tool", TOOL_NAME, action, args, False, err.code, err.message, elapsed_ms=err.elapsed_ms)
        return err
    src = ok_path

    # 第二个及之后的文件也要逐个过安全校验，不能只查第一个
    others: list[Path] = []
    for item in _split_paths(other_paths):
        op, oerr = validate_input_file(item, tool, action)
        if oerr is not None:
            oerr.message = f"附属文件校验失败：{oerr.message}"
            oerr.elapsed_ms = int((time.perf_counter() - t0) * 1000)
            log_call(session_id, "tool", TOOL_NAME, action, args, False, oerr.code, oerr.message,
                     elapsed_ms=oerr.elapsed_ms)
            return oerr
        others.append(op)

    try:
        if action == "merge":
            paths = [src] + others
            if len(paths) < 2:
                r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                                    "合并至少需要两份文档：除了当前选中的文件，还要在 other_paths 里给出其它文件",
                                    elapsed_ms=int((time.perf_counter() - t0) * 1000))
                log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                         elapsed_ms=r.elapsed_ms)
                return r
            if len(paths) > MAX_MERGE_FILES:
                r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                                    f"一次最多合并 {MAX_MERGE_FILES} 份，当前 {len(paths)} 份",
                                    elapsed_ms=int((time.perf_counter() - t0) * 1000))
                log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                         elapsed_ms=r.elapsed_ms)
                return r
            # 去重：同一个文件传两次没意义
            uniq: list[Path] = []
            for p in paths:
                if all(p.resolve() != q.resolve() for q in uniq):
                    uniq.append(p)
            out = build_output_path(src, "merge", ".docx")
            with InputGuard(src):
                r = _do_merge(uniq, out, add_page_break, mark_source)

        elif action == "split":
            with InputGuard(src):
                r = _do_split(src, level, anchor)

        elif action == "extract_tables":
            if src.suffix.lower() == ".pdf":
                r = ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT,
                                    "PDF 里没有可提取的表格对象，请先转成 Word 再提取",
                                    elapsed_ms=int((time.perf_counter() - t0) * 1000))
                log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                         elapsed_ms=r.elapsed_ms)
                return r
            out = build_output_path(src, "tables", ".xlsx")
            with InputGuard(src):
                r = _do_extract_tables(src, out)

        else:  # diff
            if not others:
                r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                                    "对比需要两份文档：请在 other_paths 里给出另一份（新版）文件",
                                    elapsed_ms=int((time.perf_counter() - t0) * 1000))
                log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                         elapsed_ms=r.elapsed_ms)
                return r
            out = build_output_path(src, "diff", ".docx")
            with InputGuard(src):
                r = _do_diff(src, others[0], out)

        log_call(session_id, "tool", TOOL_NAME, action, args, r.ok, r.code, r.message,
                 output_file=r.output_file, elapsed_ms=r.elapsed_ms,
                 extra={"count": len(r.changes)})
        return r

    except Exception as e:
        msg = f"结构操作失败：{type(e).__name__}: {e}"
        r = ToolResult.fail(tool, action, CODE_INTERNAL_ERROR, msg,
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                 elapsed_ms=r.elapsed_ms)
        return r


def _lc_entry(path: str, action: str, other_paths: str = "", level: int = 1,
              anchor: str = "", add_page_break: bool = True, mark_source: bool = False) -> str:
    return json.dumps(
        run(path, action, other_paths, level, anchor, add_page_break, mark_source).to_dict(),
        ensure_ascii=False,
    )


def build_tool():
    from langchain_core.tools import StructuredTool

    return StructuredTool.from_function(
        func=_lc_entry, name=TOOL_NAME, description=TOOL_DESC, args_schema=StructureArgs,
    )
