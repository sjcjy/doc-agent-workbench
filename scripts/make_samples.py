# -*- coding: utf-8 -*-
"""生成两份结构差异明显的样例 DOCX，供演示与测试使用。

样例一：带封面标题、多级标题、项目符号、表格的「项目周报」
样例二：纯 Title + 二级标题 + 长正文段落的「实验报告」

两者结构不同，正好用来验证「换文件也照样跑」的参数化要求。
运行：python scripts/make_samples.py
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from core import config
from core.tools import _docxkit as dk


def build_weekly_report(path: Path) -> Path:
    doc = Document()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("智能文档助手项目周报")
    dk.set_run_font(r, cn_font="黑体", size_pt=20, bold=True)

    dk.add_heading(doc, "一、本周进展", 1)
    for t in [
        "完成 DOCX 解析模块的开发，支持段落、表格与样式的读取。",
        "接入大模型接口，实现任务意图识别与工具选择。",
        "完成三份样例文档的格式适配测试，覆盖率达到 85%。",
    ]:
        bp = doc.add_paragraph(t)
        dk.set_line_spacing(bp, 1.5)
        dk.set_first_line_indent(bp, chars=2)

    dk.add_heading(doc, "1.1 关键数据", 2)
    table = doc.add_table(rows=4, cols=3)
    table.style = dk.resolve_style(doc, "Table Grid", "网格型") or table.style
    heads = ["指标", "本周", "上周"]
    vals = [["文档处理量", "128", "96"], ["平均耗时(秒)", "3.2", "4.7"], ["成功率", "98.4%", "95.1%"]]
    for i, h in enumerate(heads):
        cell = table.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
    for ri, row in enumerate(vals, start=1):
        for ci, v in enumerate(row):
            table.rows[ri].cells[ci].text = v

    dk.add_heading(doc, "二、存在问题", 1)
    for t in ["跨 run 文本匹配偶发失败，需在段落级做合并处理。", "PDF 导出依赖外部软件，打包后体积偏大。"]:
        bp = doc.add_paragraph(t)
        dk.set_first_line_indent(bp, chars=2)

    dk.add_heading(doc, "三、下周计划", 1)
    for t in ["完善异常处理与错误提示文案。", "补充不少于 12 条测试记录。", "整理演示脚本与截图证据。"]:
        bp = doc.add_paragraph(t)
        dk.set_first_line_indent(bp, chars=2)

    bp = doc.add_paragraph("联系邮箱：project@example.com")
    dk.set_first_line_indent(bp, chars=2)
    doc.save(str(path))
    return path


def build_lab_report(path: Path) -> Path:
    doc = Document()
    try:
        title = doc.add_paragraph("基于 LangChain 的文档处理实验报告", style="Title")
    except KeyError:
        title = doc.add_paragraph()
        dk.set_run_font(title.add_run("基于 LangChain 的文档处理实验报告"), size_pt=22, bold=True)

    intro = doc.add_paragraph(
        "本次实验面向日常办公场景中的文档自动化处理需求，尝试使用大语言模型理解用户的自然语言指令，"
        "再由确定性的程序工具完成具体的文件操作。实验重点关注三个问题：指令到工具调用的映射是否准确、"
        "文件操作过程是否安全可靠、以及最终产出的文档能否直接在 Word 中正常打开编辑。"
    )
    dk.set_line_spacing(intro, 1.5)
    dk.set_first_line_indent(intro, chars=2)

    dk.add_heading(doc, "实验方法", 2)
    body = doc.add_paragraph(
        "实验采用分层架构。最上层是交互界面，负责收集用户上传的文档与自然语言任务描述；"
        "中间层是智能体，负责理解意图并选择恰当的工具；最下层是工具集合，"
        "每个工具只完成一类确定性操作。这样的划分使得模型只承担理解与决策，"
        "所有文件读写都由可测试、可复现的程序代码完成。联系邮箱：lab@example.com"
    )
    dk.set_line_spacing(body, 1.5)
    dk.set_first_line_indent(body, chars=2)

    dk.add_heading(doc, "结果与分析", 2)
    analysis = doc.add_paragraph(
        "在二十组自然语言指令的测试中，系统成功完成十八组，失败的两组均源于指令本身表述模糊，"
        "而非工具执行错误。所有输出文件均能在 Word 中正常打开，原文件在操作前后哈希值保持一致。"
    )
    dk.set_line_spacing(analysis, 1.5)
    dk.set_first_line_indent(analysis, chars=2)

    dk.add_heading(doc, "结论", 2)
    concl = doc.add_paragraph(
        "分层架构在保证灵活性的同时兼顾了可靠性，模型负责理解、工具负责执行的职责划分是有效的。"
    )
    dk.set_first_line_indent(concl, chars=2)
    doc.save(str(path))
    return path


def main() -> None:
    config.ensure_dir(config.SAMPLE_DIR)
    a = build_weekly_report(config.SAMPLE_DIR / "样例1_项目周报.docx")
    b = build_lab_report(config.SAMPLE_DIR / "样例2_实验报告.docx")
    for f in (a, b):
        doc = Document(str(f))
        print(f"[OK] {f.name}  段落={len(doc.paragraphs)} 表格={len(doc.tables)} 字数≈{dk.word_count(doc)}")


if __name__ == "__main__":
    main()
