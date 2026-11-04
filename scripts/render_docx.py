# -*- coding: utf-8 -*-
"""把 docx 用 Word 导出 PDF 再渲染成 PNG，用于肉眼验收（不靠猜）。"""
import sys
from pathlib import Path

import pythoncom
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent


def docx_to_pdf(src: Path, out_pdf: Path) -> None:
    src, out_pdf = src.resolve(), out_pdf.resolve()
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.SaveAs2(str(out_pdf), FileFormat=17)  # wdFormatPDF
        finally:
            doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def pdf_to_png(pdf: Path, out_dir: Path, prefix: str, dpi: int = 110) -> list[Path]:
    import pymupdf as fitz
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    with fitz.open(str(pdf)) as doc:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi)
            p = out_dir / f"{prefix}_p{i}.png"
            pix.save(str(p))
            paths.append(p)
    return paths


def metrics(pdf: Path) -> dict:
    import pymupdf as fitz
    over = 0
    max_right = 0.0
    with fitz.open(str(pdf)) as doc:
        page = doc[0]
        W = page.rect.width
        # 版心右边界（按 2.5cm 页边距估算）
        limit = W - 2.5 / 2.54 * 72
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                x1 = line["bbox"][2]
                max_right = max(max_right, x1)
                if x1 > limit + 1:
                    over += 1
    return {"pages": len(fitz.open(str(pdf))), "page_w": round(W, 1),
            "limit": round(limit, 1), "max_right": round(max_right, 1), "over": over}


def main():
    src = Path(sys.argv[1])
    prefix = sys.argv[2] if len(sys.argv) > 2 else src.stem
    out_pdf = ROOT / "evidence" / "preview" / f"{prefix}.pdf"
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    docx_to_pdf(src, out_pdf)
    paths = pdf_to_png(out_pdf, ROOT / "evidence" / "preview", prefix)
    print("pdf:", out_pdf)
    for p in paths:
        print("png:", p)
    print("metrics:", metrics(out_pdf))


if __name__ == "__main__":
    main()
