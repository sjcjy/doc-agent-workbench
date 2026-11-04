from __future__ import annotations

"""工具五：文档格式转换（convert）。

支持：
- docx → txt    纯文本内容提取（含表格）
- docx → md     Markdown（标题/列表/表格都还原成 Markdown 语法）
- docx → pdf    多后端自动探测：LibreOffice 无头模式 → Word COM → 纯 Python 渲染
- txt  → docx
- md   → docx

关于 PDF 的态度：装了哪个用哪个，都不行就明确返回失败并给出安装建议，
绝不生成一个空 PDF 冒充成功（计划书「如实返回，不伪造成功」）。
"""

import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from pydantic import BaseModel, Field
from typing_extensions import Literal

from core import config
from core.ctx import get_session_id
from core.logger import log_call
from core.models import (
    CODE_FILE_CORRUPTED,
    CODE_INTERNAL_ERROR,
    CODE_INVALID_INPUT,
    CODE_RENDER_FAILED,
    CODE_UNSUPPORTED_FORMAT,
    ToolResult,
)
from core.safety import InputGuard, build_output_path, validate_input_file
from core.tools import _docxkit as dk
from core.tools._mdrender import render_markdown_to_docx

TOOL_NAME = "convert_document"
TOOL_DESC = (
    "文档格式转换，输出为新文件。"
    "target=txt 把 DOCX 转成纯文本；target=md 转成 Markdown（保留标题层级与表格）；"
    "target=pdf 转成 PDF（自动探测 LibreOffice / Word / 内置渲染）；"
    "target=docx 时按源文件区分：源为 TXT/MD 则生成 Word 文档，"
    "源为 PDF 则把 PDF 转成 Word（尽量保留排版；扫描件无文本层会失败并说明原因）。"
)

ConvertTarget = Literal["txt", "md", "markdown", "pdf", "docx"]

# 常见的 LibreOffice 安装位置，按平台依次探测
_SOFFICE_CANDIDATES = [
    "soffice", "libreoffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice", "/usr/bin/libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


# ---------------------------------------------------------------- 后端探测
def find_soffice() -> str | None:
    for c in _SOFFICE_CANDIDATES:
        p = shutil.which(c)
        if p:
            return p
        pc = Path(c)
        if pc.is_file():
            return str(pc)
    return None


def has_word_com() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def pdf_backends() -> dict[str, bool]:
    """返回各 PDF 后端的可用情况，界面可以据此给出安装提示。"""
    return {
        "rebuild": True,
        "word_com": has_word_com(),
        "pdf2docx": _has_pdf2docx(),
        "text_only": True,
    }


def _has_pdf2docx() -> bool:
    try:
        import pdf2docx  # noqa: F401
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 提取
def docx_to_text(doc) -> str:
    """DOCX → 纯文本。标题单独成行，表格转成制表符分隔。"""
    lines: list[str] = []
    for p in doc.paragraphs:
        t = dk.para_text(p).strip()
        if not t:
            continue
        lv = dk.heading_level(p)
        lines.append(("#" * lv + " " + t) if lv else t)
    for tb in doc.tables:
        lines.append("")
        for row in tb.rows:
            lines.append("\t".join(c.text.strip() for c in row.cells))
    return "\n".join(lines).strip() + "\n"


def docx_to_markdown(doc) -> str:
    """DOCX → Markdown。尽量还原标题、列表、加粗与表格。"""
    out: list[str] = []
    for p in doc.paragraphs:
        t = dk.para_text(p).strip()
        if not t:
            continue
        lv = dk.heading_level(p)
        style = (p.style.name if p.style is not None else "") or ""

        body_md = ""
        # 逐个 run 处理粗体
        for r in p.runs:
            seg = r.text.replace("\n", " ")
            if not seg:
                continue
            if r.bold and len(seg.strip()) <= 40:
                body_md += f"**{seg.strip()}**"
            else:
                body_md += seg
        body_md = body_md.strip() or t

        if lv:
            out.append(f"{'#' * lv} {body_md}")
        elif "List Bullet" in style or "项目符号" in style:
            out.append(f"- {body_md}")
        elif "List Number" in style or "编号" in style:
            out.append(f"1. {body_md}")
        elif "Quote" in style or "引用" in style:
            out.append(f"> {body_md}")
        else:
            out.append(body_md)
        out.append("")

    for tb in doc.tables:
        rows = [[" ".join(c.text.split()) for c in row.cells] for row in tb.rows]
        if not rows:
            continue
        out.append("| " + " | ".join(rows[0]) + " |")
        out.append("| " + " | ".join("---" for _ in rows[0]) + " |")
        for r in rows[1:]:
            out.append("| " + " | ".join(r) + " |")
        out.append("")

    text = "\n".join(out)
    return re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"


# ---------------------------------------------------------------- PDF 三种后端
def _pdf_via_libreoffice(src: Path, out_pdf: Path, timeout: int = 120) -> tuple[bool, str]:
    exe = find_soffice()
    if not exe:
        return False, "未检测到 LibreOffice"
    try:
        cmd = [exe, "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", str(out_pdf.parent), str(src)]
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout, shell=False)
        produced = src.with_suffix(".pdf")
        produced = out_pdf.parent / (src.stem + ".pdf")
        if proc.returncode == 0 and produced.exists():
            if produced != out_pdf:
                produced.replace(out_pdf)
            return True, "LibreOffice 无头模式"
        err = (proc.stderr or b"").decode("utf-8", "ignore")[:200]
        return False, f"LibreOffice 转换失败(returncode={proc.returncode}) {err}"
    except subprocess.TimeoutExpired:
        return False, "LibreOffice 转换超时"
    except Exception as e:
        return False, f"LibreOffice 调用异常：{type(e).__name__}: {e}"


def _pdf_via_word(src: Path, out_pdf: Path, timeout: int = 120) -> tuple[bool, str]:
    if not has_word_com():
        return False, "未安装 pywin32，无法调用 Word"
    try:
        import win32com.client as win32
        import pythoncom
        pythoncom.CoInitialize()
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(src), ReadOnly=True)
            # 0 = wdFormatPDF
            doc.SaveAs(str(out_pdf), FileFormat=17)
            doc.Close(False)
        finally:
            word.Quit()
            pythoncom.CoUninitialize()
        return out_pdf.exists(), ("Word COM" if out_pdf.exists() else "Word COM 未生成文件")
    except Exception as e:
        return False, f"Word COM 调用异常：{type(e).__name__}: {e}"


def _pdf_via_builtin(src: Path, out_pdf: Path) -> tuple[bool, str]:
    """纯 Python 渲染：不依赖任何外部软件，保证在最坏情况下也能出 PDF。

    用 reportlab 内置的 Adobe CID 中文字体 STSong-Light，无需系统字体文件。
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError as e:
        return False, f"缺少渲染依赖 reportlab：{e}"

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        cn_font = "STSong-Light"
    except Exception:
        cn_font = "Helvetica"

    try:
        doc = dk.open_docx(src)
        styles = getSampleStyleSheet()
        base = ParagraphStyle("cn", parent=styles["Normal"], fontName=cn_font, fontSize=11,
                              leading=18, alignment=TA_JUSTIFY, spaceAfter=6)
        h1 = ParagraphStyle("h1", parent=base, fontSize=18, leading=26, alignment=TA_CENTER, spaceAfter=12)
        h2 = ParagraphStyle("h2", parent=base, fontSize=15, leading=22, spaceAfter=8)
        h3 = ParagraphStyle("h3", parent=base, fontSize=13, leading=20, spaceAfter=6)

        story = []
        for p in doc.paragraphs:
            t = dk.para_text(p).strip()
            if not t:
                continue
            lv = dk.heading_level(p)
            safe = (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
            style = {1: h1, 2: h2, 3: h3}.get(lv)
            if style is None and lv:
                style = h3
            if style is None:
                style = base
            story.append(Paragraph(safe, style))

        for tb in doc.tables:
            data = [[" ".join(c.text.split()) for c in row.cells] for row in tb.rows]
            if not data:
                continue
            table = Table(data, hAlign="LEFT")
            table.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), cn_font),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f0f0")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            story.append(Spacer(1, 8))
            story.append(table)

        if not story:
            story.append(Paragraph("（文档为空）", base))

        SimpleDocTemplate(
            str(out_pdf), pagesize=A4,
            leftMargin=2.2 * cm, rightMargin=2.2 * cm,
            topMargin=2.5 * cm, bottomMargin=2.5 * cm,
            title=src.stem,
        ).build(story)
        return out_pdf.exists(), "内置渲染器（reportlab）"
    except Exception as e:
        return False, f"内置渲染失败：{type(e).__name__}: {e}"


# ---------------------------------------------------------------- PDF → DOCX
def _repair_layout(docx_path: Path) -> str:
    """pdf2docx 产物做一次版式修复（列宽/行高/行距），返回人话说明。

    失败不影响转换结果本身——修复是"加分项"，不能让主流程挂掉。
    """
    try:
        from core.tools._docxrepair import repair_docx_file

        stats = repair_docx_file(docx_path)
    except Exception:
        return ""
    parts = []
    if stats.get("tables"):
        parts.append(f"{stats['tables']} 个表格已改为自适应行高")
    if stats.get("merged"):
        parts.append(f"{stats['merged']} 处被拆散的单元格已合并")
    if stats.get("cell_paras"):
        parts.append(f"{stats['cell_paras']} 处单元格定位残留已清除")
    if stats.get("paragraphs"):
        parts.append(f"{stats['paragraphs']} 处行距/缩进已修正")
    if stats.get("scaled"):
        parts.append(f"{stats['scaled']} 个表格宽度已按版心校正")
    if not parts:
        return ""
    return "；已自动修正排版：" + "、".join(parts)


def _docx_via_word(src: Path, out_docx: Path) -> tuple[bool, str]:
    """用本机 Microsoft Word 自带的 PDF 转换（PDF Reflow）。

    实测这是保真度最高的路径：表格能识别成真正的 Word 表格，而不是靠绝对
    坐标"画"出来的假表格。仅 Windows + 装了 Word 时可用，否则调用方回落。
    用 DispatchEx 另起一个 Word 进程，避免影响用户已经打开着的 Word 文档。
    """
    if sys.platform != "win32":
        return False, "非 Windows 环境"
    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except ImportError:
        return False, "未安装 pywin32"

    # Word 会按自己的进程目录解析相对路径，必须传绝对路径
    src_abs = Path(src).resolve()
    out_abs = Path(out_docx).resolve()
    if not src_abs.exists():
        return False, "源文件不存在"

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.Options.ConfirmConversions = False
        except Exception:
            pass
        doc = word.Documents.Open(str(src_abs), ConfirmConversions=False,
                                  ReadOnly=False, AddToRecentFiles=False)
        try:
            doc.SaveAs2(str(out_abs), FileFormat=12)  # 12 = wdFormatXMLDocument
        finally:
            doc.Close(False)
        ok = out_abs.exists() and out_abs.stat().st_size > 0
        return ok, ("Microsoft Word（PDF Reflow，保真度最高）" if ok else "Word 未产出文件")
    except Exception as e:
        return False, f"Word 转换失败：{type(e).__name__}: {e}"
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _docx_from_pdf(src: Path, out_docx: Path) -> tuple[bool, str, dict]:
    """PDF → DOCX，按"先结构、后兜底"的顺序尝试。

    方案一 结构重建（默认）：从 PDF 里读出网格线与文字坐标，重建成真正的 Word
        表格与段落。产物可编辑、可自动换行，字号/页边距怎么改都不会散。
        —— 这是唯一能真正解决"转换后排版乱"的路径。
    方案二 Microsoft Word（PDF Reflow）：装 Office 才有。保真度高，但内部是上千个
        浮动文本框，位置靠锚点维持，一旦改动就整体错位，所以只作兜底。
    方案三 pdf2docx：纯 Python，把绝对坐标写成段落缩进，复杂表格识别弱。
    方案四 PyMuPDF 纯文本：只保文字，最后兜底。

    扫描件（图片型 PDF）没有文本层，全部失败时如实报告，不伪造内容。
    """
    first_err = ""

    # 方案一：结构重建
    try:
        from . import _pdfrebuild

        ok, note, stats = _pdfrebuild.convert(src, out_docx)
        if ok and out_docx.exists() and out_docx.stat().st_size > 0:
            return True, note, stats
        first_err = note
    except Exception as e:
        first_err = f"结构重建失败：{type(e).__name__}: {e}"

    # 方案二：Microsoft Word（PDF Reflow）
    ok, note = _docx_via_word(src, out_docx)
    if ok:
        return True, note + "（该路径为浮动文本框排版，建议改动后再用「格式」工具统一版式）", {}
    first_err += f"；{note}"

    # 方案三：pdf2docx
    try:
        from pdf2docx import Converter  # type: ignore

        cv = Converter(str(src))
        try:
            cv.convert(str(out_docx))
        finally:
            cv.close()
        if out_docx.exists() and out_docx.stat().st_size > 0:
            note = _repair_layout(out_docx)
            return True, "pdf2docx（保留排版与表格）" + note, {}
        first_err += "；pdf2docx 未产出文件"
    except ImportError:
        first_err += "；未安装 pdf2docx"
    except Exception as e:
        first_err += f"；pdf2docx 失败：{type(e).__name__}: {e}"

    # 方案四：PyMuPDF 纯文本重建
    try:
        try:
            import pymupdf as fitz  # PyMuPDF >= 1.24 的新入口
        except ImportError:
            import fitz  # 旧版本兼容
        from docx import Document

        doc = Document()
        total_chars = 0
        with fitz.open(str(src)) as pdf:
            n_pages = len(pdf)
            for i, page in enumerate(pdf):
                text = page.get_text("text")
                if text.strip():
                    for line in text.splitlines():
                        t = line.strip()
                        if not t:
                            continue
                        doc.add_paragraph(t)
                        total_chars += len(t)
                    if i < n_pages - 1:
                        doc.add_page_break()
        if total_chars == 0:
            return False, (
                f"没能从 PDF 提取到文字（{first_err}）。"
                "这份 PDF 可能是扫描件/图片型，需要先 OCR 才能转换。"
            ), {}
        doc.save(str(out_docx))
        return True, f"PyMuPDF 文本重建（{total_chars} 字符，仅保留文字，排版已简化）", {}
    except ImportError:
        return False, (
            f"PDF 转换失败（{first_err}）。缺少 PDF 解析依赖，"
            "请在环境中执行：pip install pdf2docx"
        ), {}
    except Exception as e:
        return False, f"PDF 转换失败（{first_err}）；文本重建也失败：{type(e).__name__}: {e}", {}


# ---------------------------------------------------------------- 主入口
class ConvertArgs(BaseModel):
    path: str = Field(description="源文件绝对路径（doc 类为 DOCX，生成类可为 txt/md）")
    target: ConvertTarget = Field(description="目标格式：txt / md / pdf / docx")
    pdf_backend: str = Field(default="auto", description="pdf 后端：auto / libreoffice / word / builtin")
    font: str = Field(default="", description="生成 DOCX 时的中文字体")
    font_size_pt: float = Field(default=config.DEFAULT_BODY_SIZE_PT, description="生成 DOCX 时的字号")


def run(path: str, target: str, pdf_backend: str = "auto", font: str = "",
        font_size_pt: float = config.DEFAULT_BODY_SIZE_PT, session_id: str = "") -> ToolResult:
    session_id = session_id or get_session_id()
    t0 = time.perf_counter()
    tool, action = "convert", f"to_{target}"
    args = {"path": path, "target": target, "pdf_backend": pdf_backend}

    ok_path, err = validate_input_file(path, tool, action)
    if err is not None:
        err.elapsed_ms = int((time.perf_counter() - t0) * 1000)
        log_call(session_id, "tool", TOOL_NAME, action, args, False, err.code, err.message, elapsed_ms=err.elapsed_ms)
        return err
    src = ok_path
    ext = src.suffix.lower()

    with InputGuard(src) as guard:
        try:
            # ---------------- 源为 DOCX ----------------
            if ext == ".docx":
                doc = dk.open_docx(src)

                if target in ("txt",):
                    out = build_output_path(src, action, ".txt")
                    text = docx_to_text(doc)
                    out.write_text(text, encoding="utf-8")
                    r = ToolResult.ok_result(tool, action, f"已转为 TXT：{len(text)} 字符",
                                             output_file=str(out),
                                             data={"chars": len(text), "source": str(src)},
                                             elapsed_ms=int((time.perf_counter() - t0) * 1000))

                elif target in ("md", "markdown"):
                    out = build_output_path(src, action, ".md")
                    text = docx_to_markdown(doc)
                    out.write_text(text, encoding="utf-8")
                    r = ToolResult.ok_result(tool, action, f"已转为 Markdown：{len(text)} 字符",
                                             output_file=str(out),
                                             data={"chars": len(text), "source": str(src)},
                                             elapsed_ms=int((time.perf_counter() - t0) * 1000))

                elif target == "pdf":
                    out = build_output_path(src, action, ".pdf")
                    order = {
                        "auto": ["libreoffice", "word", "builtin"],
                        "libreoffice": ["libreoffice"],
                        "word": ["word"],
                        "builtin": ["builtin"],
                    }.get(str(pdf_backend).lower(), ["libreoffice", "word", "builtin"])

                    tried: list[dict] = []
                    used = ""
                    for backend in order:
                        if backend == "libreoffice":
                            okb, info = _pdf_via_libreoffice(src, out)
                        elif backend == "word":
                            okb, info = _pdf_via_word(src, out)
                        else:
                            okb, info = _pdf_via_builtin(src, out)
                        tried.append({"backend": backend, "ok": okb, "info": info})
                        if okb:
                            used = info
                            break

                    if used:
                        size = out.stat().st_size if out.exists() else 0
                        r = ToolResult.ok_result(
                            tool, action, f"已转为 PDF（{used}），大小 {size / 1024:.0f} KB",
                            output_file=str(out),
                            data={"backend": used, "bytes": size, "tried": tried, "backends": pdf_backends()},
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    else:
                        detail = "；".join(f"{t['backend']}={t['info']}" for t in tried)
                        r = ToolResult.fail(
                            tool, action, CODE_RENDER_FAILED,
                            f"PDF 转换失败，所有后端均未产出文件：{detail}。"
                            f"建议安装 LibreOffice（推荐，保真度最高）：https://www.libreoffice.org/download/",
                            data={"tried": tried, "backends": pdf_backends()},
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
                        log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                        return r

                else:
                    r = ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT,
                                        f"DOCX 不能转换成 {target}（可选 txt / md / pdf）",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r

            # ---------------- 源为 PDF ----------------
            elif ext == ".pdf":
                if target == "pdf":
                    r = ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT,
                                        "源文件已经是 PDF，无需再次转换",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                             elapsed_ms=r.elapsed_ms)
                    return r
                if target != "docx":
                    r = ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT,
                                        f"PDF 目前只能转成 DOCX，不能转成 {target}。建议先转成 DOCX 再继续处理。",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                             elapsed_ms=r.elapsed_ms)
                    return r

                out = build_output_path(src, action, ".docx")
                okb, info, extra = _docx_from_pdf(src, out)
                if okb:
                    data = {"source": str(src), "method": info}
                    if extra:
                        data["stats"] = extra
                    r = ToolResult.ok_result(tool, action, f"已由 PDF 生成 DOCX：{info}",
                                             output_file=str(out),
                                             data=data,
                                             elapsed_ms=int((time.perf_counter() - t0) * 1000))
                else:
                    r = ToolResult.fail(tool, action, CODE_RENDER_FAILED, info,
                                        data={"source": str(src)},
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message,
                             elapsed_ms=r.elapsed_ms)
                    return r

            # ---------------- 源为 TXT / MD ----------------
            elif ext in config.SUPPORTED_TEXT_EXT:
                if target != "docx":
                    r = ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT,
                                        f"{ext} 只能转成 docx，不能转成 {target}",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                try:
                    raw = src.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    raw = src.read_text(encoding="gbk", errors="ignore")
                except Exception as e:
                    r = ToolResult.fail(tool, action, CODE_FILE_CORRUPTED, f"无法读取文本文件：{type(e).__name__}: {e}",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                if not raw.strip():
                    r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "文本文件内容为空，无法生成 DOCX",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                out = build_output_path(src, action, ".docx")
                if ext == ".txt":
                    # TXT 没有 Markdown 标记，把非空行当作段落
                    md = "\n\n".join(line.strip() for line in raw.splitlines() if line.strip())
                else:
                    md = raw
                render_markdown_to_docx(md, title=src.stem, output_path=out,
                                        body_font=font or None, body_size_pt=font_size_pt)
                r = ToolResult.ok_result(tool, action, f"已由 {ext} 生成 DOCX",
                                         output_file=str(out), data={"source": str(src)},
                                         elapsed_ms=int((time.perf_counter() - t0) * 1000))
            else:
                r = ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT, f"不支持的源文件类型：{ext}",
                                    elapsed_ms=int((time.perf_counter() - t0) * 1000))
                log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                return r

            return r

        except Exception as e:
            msg = f"转换失败：{type(e).__name__}: {e}"
            r = ToolResult.fail(tool, action, CODE_INTERNAL_ERROR, msg,
                                elapsed_ms=int((time.perf_counter() - t0) * 1000))
            log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
            return r

    if guard.violated:
        return ToolResult.fail(tool, action, "SAFETY_BLOCKED", "源文件在处理过程中被改动，已中止",
                               elapsed_ms=int((time.perf_counter() - t0) * 1000))
    log_call(session_id, "tool", TOOL_NAME, action, args, r.ok, r.code, r.message,
             output_file=r.output_file, elapsed_ms=r.elapsed_ms, extra={"guard": guard.report()})
    return r


def _lc_entry(path: str, target: str, pdf_backend: str = "auto", font: str = "",
              font_size_pt: float = config.DEFAULT_BODY_SIZE_PT) -> str:
    return json.dumps(run(path, target, pdf_backend, font, font_size_pt).to_dict(), ensure_ascii=False)


def build_tool():
    from langchain_core.tools import StructuredTool
    return StructuredTool.from_function(
        func=_lc_entry, name=TOOL_NAME, description=TOOL_DESC, args_schema=ConvertArgs,
    )
