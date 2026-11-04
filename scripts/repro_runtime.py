# -*- coding: utf-8 -*-
"""复现「同一份 PDF，脚本里转得好、跑起来转得差」的问题。

分两种执行环境调用 Word 后端：
  1) 主线程（我之前的验收脚本就是这么跑的）
  2) 子线程（Streamlit 的脚本是跑在 ScriptRunner 子线程里的，FastAPI 是线程池）
"""
import sys
import threading
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.tools import convert_tool  # noqa: E402

SRC = ROOT / "samples" / "_pdf_probe.pdf"


def ensure_src() -> Path:
    if SRC.exists():
        return SRC
    # 现造一份两页的 PDF（有文本层），跟真实场景同类
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    doc = fitz.open()
    for i in range(2):
        page = doc.new_page()
        text = "\n".join(
            [f"第 {i + 1} 页 段落测试"]
            + [f"{i + 1}.{j} 这一行用来检查换行与行距是否正常。" for j in range(1, 12)]
        )
        page.insert_textbox(fitz.Rect(60, 60, 500, 760), text, fontname="china-s", fontsize=11)
    SRC.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(SRC))
    doc.close()
    return SRC


def probe(tag: str) -> None:
    src = ensure_src()
    out = ROOT / "data" / "outputs" / f"_probe_{tag}.docx"
    if out.exists():
        out.unlink()
    try:
        ok, note = convert_tool._docx_via_word(src, out)
        size = out.stat().st_size if out.exists() else 0
        print(f"[{tag}] ok={ok} note={note} size={size}")
    except Exception:
        print(f"[{tag}] EXCEPTION\n{traceback.format_exc()}")


def main() -> None:
    probe("main")

    box = {}

    def worker():
        try:
            import pythoncom  # noqa: F401
        except Exception as e:
            print("import pythoncom failed:", e)
        probe("thread")

    t = threading.Thread(target=worker, name="ScriptRunner")
    t.start()
    t.join()
    print("box:", box)


if __name__ == "__main__":
    main()
