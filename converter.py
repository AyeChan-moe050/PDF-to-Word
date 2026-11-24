import argparse
import io
from pathlib import Path
from typing import Callable, Optional
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_run_font(run, font_name: str = "Pyidaungsu", font_size: int = 11) -> None:
    """Apply font cleanly for both Latin (English) and complex/Myanmar scripts."""
    run.font.name = font_name
    run.font.size = Pt(font_size)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)


def ocr_page(page: fitz.Page, dpi: int = 300, lang: str = "mya+eng") -> str:
    """Render a PDF page to an image and run Tesseract OCR for Myanmar + English."""
    pix = page.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang=lang)


def convert_pdf_bytes_to_docx_bytes(
    pdf_bytes: bytes,
    font_name: str = "Pyidaungsu",
    font_size: int = 11,
    ocr_lang: str = "mya+eng",
    force_ocr: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> bytes:
    """Convert PDF bytes into Word (.docx) bytes in memory for Streamlit."""
    doc = Document()

    normal_style = doc.styles["Normal"]
    normal_style.font.name = font_name
    normal_style.font.size = Pt(font_size)

    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(pdf_doc)

    try:
        for page_index in range(total_pages):
            page = pdf_doc[page_index]
            text = "" if force_ocr else page.get_text("text").strip()

            if not text:
                if progress_callback:
                    progress_callback(
                        page_index + 1,
                        total_pages,
                        f"Page {page_index + 1}/{total_pages}: Running OCR ({ocr_lang})...",
                    )
                text = ocr_page(page, lang=ocr_lang).strip()
            else:
                if progress_callback:
                    progress_callback(
                        page_index + 1,
                        total_pages,
                        f"Page {page_index + 1}/{total_pages}: Extracted digital text.",
                    )

            heading = doc.add_heading(f"Page {page_index + 1}", level=2)
            for run in heading.runs:
                set_run_font(run, font_name=font_name, font_size=font_size + 2)

            lines = text.splitlines() if text else ["[No text detected on this page]"]
            for line in lines:
                p = doc.add_paragraph()
                run = p.add_run(line)
                set_run_font(run, font_name=font_name, font_size=font_size)

            if page_index < total_pages - 1:
                doc.add_page_break()

        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer.getvalue()
    finally:
        pdf_doc.close()


def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    font_name: str = "Pyidaungsu",
    font_size: int = 11,
    ocr_lang: str = "mya+eng",
    force_ocr: bool = False,
) -> None:
    """Convert a PDF file on disk to a .docx file on disk."""
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_file}")

    docx_bytes = convert_pdf_bytes_to_docx_bytes(
        pdf_bytes=pdf_file.read_bytes(),
        font_name=font_name,
        font_size=font_size,
        ocr_lang=ocr_lang,
        force_ocr=force_ocr,
    )

    out_file = Path(docx_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_bytes(docx_bytes)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert Myanmar + English PDF files to Word (.docx)."
    )
    parser.add_argument("input_pdf", help="Path to input PDF file")
    parser.add_argument("output_docx", help="Path to output .docx file")
    parser.add_argument("--font", default="Pyidaungsu", help="Word font name (default: Pyidaungsu)")
    parser.add_argument("--size", type=int, default=11, help="Font size in pt (default: 11)")
    parser.add_argument("--lang", default="mya+eng", help="Tesseract OCR languages (default: mya+eng)")
    parser.add_argument("--force-ocr", action="store_true", help="Force OCR on every page")

    args = parser.parse_args()
    convert_pdf_to_docx(
        pdf_path=args.input_pdf,
        docx_path=args.output_docx,
        font_name=args.font,
        font_size=args.size,
        ocr_lang=args.lang,
        force_ocr=args.force_ocr,
    )
    print(f"Done! Saved Word document to: {args.output_docx}")
