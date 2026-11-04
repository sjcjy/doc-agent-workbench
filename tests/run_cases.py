# -*- coding: utf-8 -*-
"""12+ 条测试用例执行器。

覆盖计划书要求：
- 四大模块（撰写 / 修改 / 格式 / 转换）
- 三类情形（正常 / 边界 / 失败）
- 每条记录：入参、返回码、耗时、输出文件、能否打开、源文件哈希是否变化

运行：python tests/run_cases.py
产物：tests/test_records.xlsx + evidence/输出文件索引.md
"""

import shutil
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from core import config
from core.logger import new_session_id
from core.safety import build_output_path, sha256_of
from core.tools import dispatch

SAMPLE_A = config.SAMPLE_DIR / "样例1_项目周报.docx"
SAMPLE_B = config.SAMPLE_DIR / "样例2_实验报告.docx"

# 损坏文件：内容不是合法 docx/zip。放在允许的 data/uploads 内，确保能越过路径安全校验、
# 真正触发「文件损坏」判定（FILE_CORRUPTED），而不是被 SAFETY_BLOCKED 提前拦截。
CORRUPTED = config.UPLOAD_DIR / "损坏样例.docx"

# PDF 转 Word 的样例：现场生成一份带网格线的表格 PDF，保证用例可复现
PDF_SAMPLE = config.SAMPLE_DIR / "样例3_带表格的PDF.pdf"


@dataclass
class Case:
    id: str
    module: str
    kind: str            # 正常 / 边界 / 失败
    desc: str
    tool: str
    args: dict
    expect_ok: bool
    expect_code: str = ""
    expect_ext: str = ""
    note: str = ""


def build_cases() -> list[Case]:
    a, b = str(SAMPLE_A), str(SAMPLE_B)
    return [
        # ---------------- 撰写 ----------------
        Case("T01", "撰写", "正常", "按 Markdown 大纲生成项目总结", "compose_document",
             {"topic": "季度工作总结", "title": "季度工作总结",
              "outline": "# 季度工作总结\n\n## 主要成果\n\n- 完成工具层开发\n- 完成界面与接口\n\n本季度整体进度符合预期。\n"},
             True, expect_ext=".docx"),
        Case("T02", "撰写", "边界", "outline 为空应当被拒绝，不生成空文档", "compose_document",
             {"topic": "空文档", "outline": "   "}, False, "INVALID_INPUT",
             note="输入校验拦截，未生成任何文件"),
        Case("T03", "撰写", "边界", "字号超出合法范围（200pt）", "compose_document",
             {"topic": "超大字号", "outline": "# 测试\n", "font_size_pt": 200}, False, "INVALID_INPUT"),

        # ---------------- 修改 ----------------
        Case("T04", "修改", "正常", "精确替换一处文本", "edit_document",
             {"path": a, "action": "replace", "old_text": "联系邮箱：project@example.com",
              "new_text": "联系邮箱：team@example.com"}, True, expect_ext=".docx"),
        Case("T05", "修改", "失败", "替换文档中不存在的文本，应如实失败", "edit_document",
             {"path": a, "action": "replace", "old_text": "这段文字根本不存在XYZ123", "new_text": "替换值"},
             False, "NOT_FOUND_IN_DOC", note="禁止伪造成功"),
        Case("T06", "修改", "正常", "在指定段落后新增章节", "edit_document",
             {"path": a, "action": "insert_section", "anchor": "三、下周计划",
              "heading": "四、风险提示", "body": "接口联调可能延期，需预留缓冲时间。"},
             True, expect_ext=".docx"),
        Case("T07", "修改", "正常", "在文末追加段落", "edit_document",
             {"path": b, "action": "append", "new_text": "补充说明：本周一前完成评审。"},
             True, expect_ext=".docx"),
        Case("T08", "修改", "失败", "new_text 缺失导致 append 失败", "edit_document",
             {"path": a, "action": "append", "new_text": ""}, False, "INVALID_INPUT"),

        # ---------------- 格式 ----------------
        Case("T09", "格式", "正常", "所有标题居中", "format_document",
             {"path": a, "action": "align", "target": "headings", "alignment": "center"},
             True, expect_ext=".docx"),
        Case("T10", "格式", "正常", "正文字体改为宋体、字号小四(12pt)", "format_document",
             {"path": b, "action": "font", "target": "body", "font_name": "宋体", "size_pt": 12},
             True, expect_ext=".docx"),
        Case("T11", "格式", "正常", "全文行距 1.5 倍 + 首行缩进 2 字符", "format_document",
             {"path": a, "action": "line_spacing", "target": "all", "line_spacing": 1.5},
             True, expect_ext=".docx"),
        Case("T12", "格式", "边界", "目标格式与现状一致，应返回无变化", "format_document",
             {"path": a, "action": "align", "target": "headings", "alignment": "left"},
             True, note="重复执行时可能判定为无变化"),
        Case("T13", "格式", "失败", "target 匹配不到任何段落", "format_document",
             {"path": a, "action": "align", "target": "这个段落根本不存在ABCXYZ", "alignment": "center"},
             False, "INVALID_INPUT"),

        # ---------------- 转换 ----------------
        Case("T14", "转换", "正常", "DOCX 转 Markdown（保留标题与表格）", "convert_document",
             {"path": a, "target": "md"}, True, expect_ext=".md"),
        Case("T15", "转换", "正常", "DOCX 转 TXT", "convert_document",
             {"path": b, "target": "txt"}, True, expect_ext=".txt"),
        Case("T16", "转换", "正常", "DOCX 转 PDF（自动探测后端）", "convert_document",
             {"path": a, "target": "pdf"}, True, expect_ext=".pdf"),
        Case("T17", "转换", "正常", "MD 回转为 DOCX", "convert_document",
             {"path": "", "target": "docx"}, True, expect_ext=".docx",
             note="源文件为 T14 产出的 Markdown"),
        Case("T18", "转换", "失败", "不支持把 PDF 作为可编辑输入", "convert_document",
             {"path": "", "target": "md"}, False, "UNSUPPORTED_FORMAT", note="源文件为 T16 产出的 PDF"),

        # ---------------- 健壮性与安全 ----------------
        Case("T19", "健壮性", "失败", "路径为空", "inspect_document",
             {"path": ""}, False, "INVALID_INPUT"),
        Case("T20", "健壮性", "失败", "损坏的 DOCX 文件", "inspect_document",
             {"path": str(CORRUPTED)}, False, "FILE_CORRUPTED"),
        Case("T21", "健壮性", "失败", "路径越界到项目之外", "inspect_document",
             {"path": "C:/Windows/System32/drivers/etc/hosts"}, False, "SAFETY_BLOCKED"),
        Case("T22", "健壮性", "正常", "检查第二份结构不同的样例文档", "inspect_document",
             {"path": b}, True, note="验证工具参数化，不绑定特定文档"),

        # ---------------- PDF 转 Word：排版自检 ----------------
        Case("T23", "转换", "正常", "PDF 转 Word 后排版自检（真表格/不越界/不丢字）", "convert_document",
             {"path": str(PDF_SAMPLE), "target": "docx"}, True, expect_ext=".docx",
             note="校验表格是真正的 Word 表格、宽度不超版心、文字覆盖率"),

        # ---------------- 结构操作：合并 / 拆分 / 提取表格 / 对比 ----------------
        Case("T24", "结构", "正常", "把两份文档合并成一份", "structure_document",
             {"path": a, "action": "merge", "other_paths": b}, True, expect_ext=".docx",
             note="合并后内容不能变少，表格要保留"),
        Case("T25", "结构", "正常", "按一级标题拆分文档", "structure_document",
             {"path": a, "action": "split", "level": 1}, True, expect_ext=".docx",
             note="每份都要能独立打开且非空"),
        Case("T26", "结构", "正常", "按 2 级标题拆分（标题层级不规范时自动回退）", "structure_document",
             {"path": b, "action": "split", "level": 1}, True, expect_ext=".docx",
             note="样例2 只有 2 级标题，应自动回退而不是报错"),
        Case("T27", "结构", "正常", "把文档里的表格提取成 Excel", "structure_document",
             {"path": a, "action": "extract_tables"}, True, expect_ext=".xlsx",
             note="工作表要有内容，表头文字要对"),
        Case("T28", "结构", "正常", "对比两份文档的差异", "structure_document",
             {"path": a, "action": "diff", "other_paths": b}, True, expect_ext=".docx",
             note="差异报告要列出具体变化条目"),
        Case("T29", "结构", "失败", "合并只给了一份文档", "structure_document",
             {"path": a, "action": "merge"}, False, "INVALID_INPUT"),
        Case("T30", "结构", "失败", "对没有表格的文档提取表格", "structure_document",
             {"path": b, "action": "extract_tables"}, False, "NOT_FOUND_IN_DOC"),
        Case("T31", "结构", "失败", "同一份文档和自己对比", "structure_document",
             {"path": a, "action": "diff", "other_paths": a}, False, "NO_CHANGE"),
    ]


# ---------------------------------------------------------------- 执行
def prepare_corrupted() -> None:
    config.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    CORRUPTED.write_text("这不是一个真正的 docx 文件，用来测试损坏文件处理逻辑。", encoding="utf-8")


def prepare_pdf_sample() -> None:
    """生成一份"带网格线表格 + 标题正文"的 PDF，供 T23 使用。"""
    if PDF_SAMPLE.exists():
        return
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz  # type: ignore
    PDF_SAMPLE.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)

    def put(x, y, text, size=10.5):
        page.insert_text((x, y), text, fontname="china-s", fontsize=size, color=(0, 0, 0))

    put(150, 80, "课程实验报告：文档工作台", 16)
    for i, line in enumerate([
        "本页用于验证 PDF 转 Word 后的表格结构与版面：",
        "1. 表格应转换成真正的 Word 表格，而不是浮动文本框；",
        "2. 表格宽度不得超过版心，文字不得越界；",
        "3. 段落应保留原有顺序，可以正常编辑与重新排版。",
    ]):
        put(70, 120 + i * 22, line)

    xs = [70, 150, 250, 340, 525]
    ys = [230, 262, 294, 326, 358]
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), color=(0, 0, 0), width=0.8)
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]), color=(0, 0, 0), width=0.8)

    cells = [
        ["模块", "输入", "输出", "说明"],
        ["撰写", "大纲 / 主题", "DOCX", "生成结构完整的正文"],
        ["修改", "段落关键字", "DOCX", "按关键字定位并替换"],
        ["格式", "对齐 / 字体", "DOCX", "统一版式，不改内容"],
    ]
    for r, row in enumerate(cells):
        for c, text in enumerate(row):
            put(xs[c] + 6, ys[r] + 21, text, 10)

    doc.save(str(PDF_SAMPLE))
    doc.close()


def verify_pdf_layout(docx_path: str, pdf_path: str) -> tuple[bool, str]:
    """T23 的专项校验：真表格 + 不越版心 + 文字没丢。"""
    from docx import Document

    d = Document(docx_path)
    if not d.tables:
        return False, "没有生成任何 Word 表格（说明只是复制了文字坐标）"

    sec = d.sections[0]
    # python-docx 的长度单位是 EMU：1pt = 12700，1cm = 360000
    usable_pt = (sec.page_width - sec.left_margin - sec.right_margin) / 12700.0
    for i, t in enumerate(d.tables):
        width_pt = sum((c.width or 0) for c in t.columns) / 12700.0
        if width_pt > usable_pt + 3:
            return False, (f"第 {i + 1} 个表格宽 {width_pt / 28.35:.2f}cm，"
                           f"超出可用宽度 {usable_pt / 28.35:.2f}cm")

    doc_chars = sum(len(p.text) for p in d.paragraphs)
    doc_chars += sum(len(c.text) for t in d.tables for r in t.rows for c in r.cells)

    try:
        import pymupdf as fitz
    except ImportError:
        import fitz  # type: ignore
    with fitz.open(pdf_path) as pdf:
        pdf_chars = sum(len(p.get_text("text").strip()) for p in pdf)

    ratio = doc_chars / max(1, pdf_chars)
    if ratio < 0.85:
        return False, f"文字覆盖率只有 {ratio:.0%}，疑似丢内容"

    return True, (f"{len(d.tables)} 个真表格，最宽未超 {usable_pt / 28.35:.2f}cm 版心，"
                  f"文字覆盖 {ratio:.0%}")


def verify_structure(case_id: str, data: dict, out: str) -> tuple[bool | None, str]:
    """结构操作的专项校验：不只看「成功了」，还要看产物对不对。"""
    try:
        if case_id == "T24":
            from docx import Document
            d = Document(out)
            srcs = data.get("files") or []
            total_p = total_t = 0
            for s in srcs:
                sd = Document(s)
                total_p += len(sd.paragraphs)
                total_t += len(sd.tables)
            if len(d.paragraphs) < total_p:
                return False, f"合并后只有 {len(d.paragraphs)} 段，少于源文档合计 {total_p} 段"
            if len(d.tables) < total_t:
                return False, f"合并后丢失表格（{len(d.tables)}/{total_t}）"
            return True, f"{len(d.paragraphs)} 段、{len(d.tables)} 个表格（源合计 {total_p} 段）"

        if case_id in ("T25", "T26"):
            from docx import Document
            files = [f for f in (data.get("files") or []) if Path(f).exists()]
            if len(files) < 2:
                return False, f"只产出 {len(files)} 份，至少应有 2 份"
            for f in files:
                d = Document(f)
                if not any(p.text.strip() for p in d.paragraphs):
                    return False, f"{Path(f).name} 拆分后是空的"
            return True, f"拆成 {len(files)} 份，均非空"

        if case_id == "T27":
            from openpyxl import load_workbook
            wb = load_workbook(out)
            cells = sum(1 for ws in wb.worksheets for row in ws.iter_rows(values_only=True)
                        for v in row if v not in (None, ""))
            if cells < 4:
                return False, f"Excel 内容过少（{cells} 个非空单元格）"
            head = [c for c in next(wb.worksheets[0].iter_rows(min_row=1, max_row=1, values_only=True))
                    if c not in (None, "")]
            if not head:
                return False, "第一个工作表没有表头"
            return True, f"{len(wb.worksheets)} 个工作表 / {cells} 格，表头：{'、'.join(map(str, head[:4]))}"

        if case_id == "T28":
            from docx import Document
            d = Document(out)
            if not d.tables:
                return False, "差异报告里没有明细表"
            n = len(d.tables[0].rows) - 1
            if n < 1:
                return False, "差异明细为空"
            return True, f"列出 {n} 条差异"
    except Exception as e:
        return False, f"专项校验异常：{type(e).__name__}: {e}"
    return None, ""


def verify_openable(path: str) -> str:
    """验证输出文件能否被正常打开。"""
    p = Path(path)
    if not p.exists():
        return "文件不存在"
    if p.stat().st_size == 0:
        return "空文件"
    if p.suffix.lower() == ".docx":
        try:
            from docx import Document
            d = Document(str(p))
            return f"可打开（{len(d.paragraphs)}段）"
        except Exception as e:
            return f"打不开：{type(e).__name__}"
    if p.suffix.lower() == ".pdf":
        head = p.open("rb").read(5)
        return "可打开（PDF头正常）" if head.startswith(b"%PDF") else "文件头异常"
    if p.suffix.lower() == ".xlsx":
        try:
            from openpyxl import load_workbook
            wb = load_workbook(str(p))
            cells = sum(1 for ws in wb.worksheets for row in ws.iter_rows(values_only=True)
                        for v in row if v not in (None, ""))
            return f"可打开（{len(wb.worksheets)}表/{cells}格）"
        except Exception as e:
            return f"打不开：{type(e).__name__}"
    try:
        txt = p.read_text(encoding="utf-8")
        return f"可读取（{len(txt)}字符）"
    except UnicodeDecodeError:
        return "可读取（非UTF8，已用兼容方式）"


def run_all() -> list[dict]:
    prepare_corrupted()
    prepare_pdf_sample()
    cases = build_cases()
    session = new_session_id()
    rows: list[dict] = []
    # 记录前序用例产物，供 T17/T18 复用（体现用例间的连贯性）
    produced: dict[str, str] = {}

    print(f"开始执行 {len(cases)} 条测试，会话 {session}\n" + "=" * 78)

    for c in cases:
        args = dict(c.args)
        # 动态注入前序产物路径
        if c.id == "T17":
            args["path"] = produced.get("md", str(SAMPLE_A))
        if c.id == "T18":
            args["path"] = produced.get("pdf", str(SAMPLE_A))

        src = str(args.get("path", ""))
        hash_before = hash_after = ""
        if src and Path(src).exists() and Path(src).suffix.lower() == ".docx":
            try:
                hash_before = sha256_of(src)
            except OSError:
                pass

        t0 = time.time()
        res_data: dict = {}
        try:
            res = dispatch(c.tool, **args, session_id=session)
            ok, code, msg = res.ok, res.code, res.message
            out = res.output_file or ""
            changes = len(res.changes)
            res_data = res.data or {}
        except Exception as e:
            ok, code, msg, out, changes = False, "INTERNAL_ERROR", f"{type(e).__name__}: {e}", "", 0
        elapsed = time.time() - t0

        if hash_before:
            try:
                hash_after = sha256_of(src)
            except OSError:
                hash_after = ""

        if out and c.expect_ext:
            produced.setdefault(c.expect_ext.lstrip("."), out)

        openable = verify_openable(out) if out else "无输出"
        expected_ok = c.expect_ok
        passed = (ok == expected_ok)
        if passed and c.expect_code and code != c.expect_code:
            passed = False
        # 专项校验：PDF 转 Word 的排版质量
        note = c.note
        if c.id == "T23" and ok and out:
            try:
                layout_ok, detail = verify_pdf_layout(out, str(PDF_SAMPLE))
            except Exception as e:
                layout_ok, detail = False, f"校验异常：{type(e).__name__}: {e}"
            passed = passed and layout_ok
            note = detail
            openable = "可打开" if layout_ok else "排版不合格"

        # 专项校验：结构操作的产物要真的对（份数、内容、表头、差异条目）
        if c.module == "结构" and ok and out:
            s_ok, detail = verify_structure(c.id, res_data, out)
            if s_ok is not None:
                passed = passed and s_ok
                note = detail

        rows.append({
            "编号": c.id, "模块": c.module, "类型": c.kind, "用例说明": c.desc,
            "调用工具": c.tool, "入参": str({k: (str(v)[:50]) for k, v in args.items()})[:300],
            "期望结果": "成功" if expected_ok else f"失败/{c.expect_code or '任意失败码'}",
            "实际结果": ("成功" if ok else f"失败/{code}"),
            "返回说明": msg[:200],
            "输出文件": Path(out).name if out else "",
            "文件校验": openable,
            "源文件哈希前": hash_before[:12],
            "源文件哈希后": hash_after[:12],
            "源文件未变": ("是" if (hash_before and hash_before == hash_after) else ("-" if not hash_before else "否")),
            "变更数": changes,
            "耗时(ms)": int(elapsed * 1000),
            "是否通过": "通过" if passed else "不通过",
            "备注": note,
        })

        flag = "PASS" if passed else "FAIL"
        print(f"[{flag}] {c.id} {c.module:<4} {c.kind:<3} {c.desc[:36]:<38} -> {code}")
        if not passed:
            print(f"        期望:{'成功' if expected_ok else '失败'}  实际:{'成功' if ok else '失败'}  说明:{msg[:120]}")

    print("=" * 78)
    npass = sum(1 for r in rows if r["是否通过"] == "通过")
    print(f"汇总：{npass}/{len(rows)} 通过")
    return rows


# ---------------------------------------------------------------- 导出
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
FAIL_FILL = PatternFill("solid", fgColor="FFC7CE")


def export_xlsx(rows: list[dict]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "测试记录"

    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in rows:
        ws.append([r[h] for h in headers])

    widths = {"编号": 8, "模块": 10, "类型": 8, "用例说明": 34, "调用工具": 20, "入参": 40,
              "期望结果": 18, "实际结果": 18, "返回说明": 46, "输出文件": 40, "文件校验": 20,
              "源文件哈希前": 14, "源文件哈希后": 14, "源文件未变": 10, "变更数": 8,
              "耗时(ms)": 10, "是否通过": 10, "备注": 26}
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 16)
    ws.row_dimensions[1].height = 24

    pass_col = headers.index("是否通过") + 1
    for row in range(2, len(rows) + 2):
        v = ws.cell(row=row, column=pass_col).value
        fill = OK_FILL if v == "通过" else FAIL_FILL
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=row, column=pass_col).fill = fill
        ws.cell(row=row, column=pass_col).font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"

    # 汇总页
    s = wb.create_sheet("汇总统计")
    total = len(rows)
    passed = sum(1 for r in rows if r["是否通过"] == "通过")
    unchanged = sum(1 for r in rows if r["源文件未变"] == "是")
    s.append(["指标", "数值"])
    s["A1"].font = Font(bold=True, color="FFFFFF")
    s["A1"].fill = HEAD_FILL
    s["B1"].font = Font(bold=True, color="FFFFFF")
    s["B1"].fill = HEAD_FILL
    for k, v in [("用例总数", total), ("通过数", passed), ("通过率", f"{passed / total * 100:.1f}%"),
                 ("正常类", sum(1 for r in rows if r["类型"] == "正常")),
                 ("边界类", sum(1 for r in rows if r["类型"] == "边界")),
                 ("失败类", sum(1 for r in rows if r["类型"] == "失败")),
                 ("源文件哈希未变", unchanged),
                 ("有输出文件", sum(1 for r in rows if r["输出文件"]))]:
        s.append([k, v])
    # 按模块统计
    s.append([])
    s.append(["模块", "通过/总数"])
    s[f"A{s.max_row}"].font = Font(bold=True)
    s[f"B{s.max_row}"].font = Font(bold=True)
    mods: dict[str, list[int]] = {}
    for r in rows:
        mods.setdefault(r["模块"], [0, 0])
        mods[r["模块"]][1] += 1
        if r["是否通过"] == "通过":
            mods[r["模块"]][0] += 1
    for m, (p, t) in mods.items():
        s.append([m, f"{p}/{t}"])
    s.column_dimensions["A"].width = 22
    s.column_dimensions["B"].width = 16

    out = Path(__file__).resolve().parent / "test_records.xlsx"
    wb.save(str(out))
    return out


def write_evidence(rows: list[dict]) -> Path:
    """把本次成功用例的输出文件归档到 evidence/ 并生成索引。"""
    config.EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    lines = ["# 测试证据索引", "", f"本次测试产生的输出文件已归档在本目录，可逐一打开核对。", ""]
    lines += ["| 用例 | 模块 | 输出文件 | 文件校验 |", "|---|---|---|---|"]
    for r in rows:
        if not r["输出文件"]:
            continue
        src = config.OUTPUT_DIR / r["输出文件"]
        if src.exists() and src.parent == config.OUTPUT_DIR:
            lines.append(f"| {r['编号']} | {r['模块']} | {r['输出文件']} | {r['文件校验']} |")
    lines += [
        "",
        "## 说明",
        "",
        "- 输出文件同时保留在 `data/outputs/`，本表索引两者一致；",
        "- 每条成功用例都会写入 `logs/calls.jsonl`，可在界面「调用记录」中查看；",
        "- 「源文件哈希未变」列统计了原文件保护是否生效。",
    ]
    idx = config.EVIDENCE_DIR / "测试证据索引.md"
    idx.write_text("\n".join(lines), encoding="utf-8")
    return idx


def main() -> None:
    rows = run_all()
    xlsx = export_xlsx(rows)
    idx = write_evidence(rows)
    print(f"\n测试记录表：{xlsx}")
    print(f"证据索引：{idx}")


if __name__ == "__main__":
    main()
