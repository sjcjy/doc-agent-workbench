"""结构工具冒烟测试：四个动作各跑一遍，打印结果与产物结构。"""
import sys
from pathlib import Path

sys.path.insert(0, ".")
from docx import Document
from core.tools import structure_tool as st
from core.tools import _docxkit as dk

S1 = Path("samples/样例1_项目周报.docx").resolve()
S2 = Path("samples/样例2_实验报告.docx").resolve()


def outline(p, n=6):
    d = Document(str(p))
    print(f"    段落 {len(d.paragraphs)} 表格 {len(d.tables)}")
    for i, para in enumerate(d.paragraphs[:n]):
        t = dk.para_text(para).strip()
        if t:
            print(f"    [{i}] H{dk.heading_level(para)} {t[:40]}")


print("=" * 60)
print("1) merge  样例1 + 样例2 + 样例2（重复应被去重）")
r = st.run(str(S1), "merge", other_paths=f"{S2}|{S2}", add_page_break=True, mark_source=True)
print("ok=", r.ok, r.code, "|", r.message)
if r.ok:
    outline(r.output_file)
    print("    file_count:", r.data["file_count"])

print("=" * 60)
print("2a) split 样例1（有 H1，按 1 级拆）")
r2 = st.run(str(S1), "split", level=1)
print("ok=", r2.ok, r2.code, "|", r2.message)
if r2.ok:
    for f in r2.data["files"]:
        print("   ", Path(f).name)
        outline(f, 3)

print("=" * 60)
print("2b) split 样例2（只有 H2，应自动回退）")
r2b = st.run(str(S2), "split", level=1)
print("ok=", r2b.ok, r2b.code, "|", r2b.message)

print("=" * 60)
print("3) extract_tables 样例1（有 1 个表）")
r3 = st.run(str(S1), "extract_tables")
print("ok=", r3.ok, r3.code, "|", r3.message)
if r3.ok:
    import openpyxl
    wb = openpyxl.load_workbook(r3.output_file)
    for ws in wb.worksheets:
        print(f"    sheet={ws.title} dims={ws.dimensions}")
        for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
            print("      ", [str(c)[:14] if c else "" for c in row])

print("=" * 60)
print("4) diff  样例1 vs 合并产物（应该能识别出新增段落）")
other = r.output_file if r.ok else str(S2)
r4 = st.run(str(S1), "diff", other_paths=other)
print("ok=", r4.ok, r4.code, "|", r4.message)
if r4.ok:
    outline(r4.output_file, 3)

print("=" * 60)
print("5) 失败路径")
print(" merge 单文件:", st.run(str(S1), "merge").code)
print(" diff 缺文件:", st.run(str(S1), "diff").code)
print(" 未知动作:", st.run(str(S1), "zip").code)
print(" 同一文件自比:", st.run(str(S1), "diff", other_paths=str(S1)).code)
