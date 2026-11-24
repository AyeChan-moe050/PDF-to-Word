import argparse
import io
import os
import re
import shutil
from pathlib import Path
from typing import Callable, List, Optional
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pytesseract
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

if not shutil.which("tesseract"):
    default_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(default_tesseract):
        pytesseract.pytesseract.tesseract_cmd = default_tesseract


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


def clean_ocr_text(raw_text: str) -> str:
    """Clean up common OCR spacing/line-break issues while keeping paragraphs."""
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in raw_text.splitlines()]
    paragraphs: List[str] = []
    current_para: List[str] = []

    for line in lines:
        if not line:
            if current_para:
                paragraphs.append(" ".join(current_para))
                current_para = []
            continue

        # Join wrapped lines within the same paragraph
        if current_para:
            prev = current_para[-1]
            # If previous line doesn't end with sentence/block punctuation, join with space
            if not re.search(r"[။!?…:]$", prev):
                current_para[-1] = f"{prev} {line}"
            else:
                paragraphs.append(" ".join(current_para))
                current_para = [line]
        else:
            current_para.append(line)

    if current_para:
        paragraphs.append(" ".join(current_para))

    cleaned = "\n\n".join(paragraphs)
    # Clean up redundant spaces around Myanmar punctuation
    cleaned = re.sub(r"\s+([၊။])", r"\1", cleaned)
    return cleaned.strip()


def ocr_page(page: fitz.Page, dpi: int = 400, lang: str = "mya+eng") -> str:
    """Render a PDF page to an image and run Tesseract OCR for Myanmar + English."""
    pix = page.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    # Grayscale + autocontrast helps sharpen text edges before OCR
    gray = ImageOps.grayscale(img)
    enhanced = ImageOps.autocontrast(gray)
    text = pytesseract.image_to_string(enhanced, lang=lang, config="--psm 6")
    return clean_ocr_text(text)


def extract_pages_text(
    pdf_bytes: bytes,
    ocr_lang: str = "mya+eng",
    force_ocr: bool = False,
    min_chars_threshold: int = 20,
    dpi: int = 400,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[str]:
    """Extract text per page from PDF bytes (with automatic/forced OCR fallback)."""
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(pdf_doc)
    pages_text: List[str] = []

    try:
        for page_index in range(total_pages):
            page = pdf_doc[page_index]
            raw_text = "" if force_ocr else page.get_text("text").strip()

            if force_ocr or len(raw_text) < min_chars_threshold:
                if progress_callback:
                    progress_callback(
                        page_index + 1,
                        total_pages,
                        f"Page {page_index + 1}/{total_pages}: Running OCR ({ocr_lang}, {dpi} DPI)...",
                    )
                text = ocr_page(page, dpi=dpi, lang=ocr_lang).strip()
            else:
                if progress_callback:
                    progress_callback(
                        page_index + 1,
                        total_pages,
                        f"Page {page_index + 1}/{total_pages}: Extracted digital text.",
                    )
                text = raw_text

            pages_text.append(text if text else "[No text detected on this page]")

        return pages_text
    finally:
        pdf_doc.close()


def pages_text_to_docx_bytes(
    pages_text: List[str],
    font_name: str = "Pyidaungsu",
    font_size: int = 11,
) -> bytes:
    """Build Word (.docx) bytes from a list of per-page strings."""
    doc = Document()

    normal_style = doc.styles["Normal"]
    normal_style.font.name = font_name
    normal_style.font.size = Pt(font_size)

    total_pages = len(pages_text)
    for page_index, text in enumerate(pages_text):
        heading = doc.add_heading(f"Page {page_index + 1}", level=2)
        for run in heading.runs:
            set_run_font(run, font_name=font_name, font_size=font_size + 2)

        for para_text in text.split("\n\n"):
            p = doc.add_paragraph()
            run = p.add_run(para_text.replace("\n", " ").strip())
            set_run_font(run, font_name=font_name, font_size=font_size)

        if page_index < total_pages - 1:
            doc.add_page_break()

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()


def convert_pdf_bytes_to_docx_bytes(
    pdf_bytes: bytes,
    font_name: str = "Pyidaungsu",
    font_size: int = 11,
    ocr_lang: str = "mya+eng",
    force_ocr: bool = False,
    dpi: int = 400,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> bytes:
    """Convert PDF bytes directly into Word (.docx) bytes."""
    pages_text = extract_pages_text(
        pdf_bytes=pdf_bytes,
        ocr_lang=ocr_lang,
        force_ocr=force_ocr,
        dpi=dpi,
        progress_callback=progress_callback,
    )
    return pages_text_to_docx_bytes(
        pages_text=pages_text,
        font_name=font_name,
        font_size=font_size,
    )


def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    font_name: str = "Pyidaungsu",
    font_size: int = 11,
    ocr_lang: str = "mya+eng",
    force_ocr: bool = False,
    dpi: int = 400,
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
        dpi=dpi,
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
    parser.add_argument("--dpi", type=int, default=400, help="OCR render DPI (default: 400)")
    parser.add_argument("--force-ocr", action="store_true", help="Force OCR on every page")

    args = parser.parse_args()
    convert_pdf_to_docx(
        pdf_path=args.input_pdf,
        docx_path=args.output_docx,
        font_name=args.font,
        font_size=args.size,
        ocr_lang=args.lang,
        force_ocr=args.force_ocr,
        dpi=args.dpi,
    )
    print(f"Done! Saved Word document to: {args.output_docx}")
