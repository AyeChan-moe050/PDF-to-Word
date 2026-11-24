import argparse
import io
import os
import shutil
import time
from pathlib import Path
from typing import Callable, List, Optional
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pytesseract
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from google import genai
from google.genai import types
from google.genai.errors import APIError

if not shutil.which("tesseract"):
    default_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(default_tesseract):
        pytesseract.pytesseract.tesseract_cmd = default_tesseract

DEFAULT_MODEL = "gemini-3.8-flash"

VISION_PROMPT = """Extract all text from this document page accurately.
Rules:
1. Preserve both Myanmar (Unicode) and English text exactly as shown in the image.
2. Keep original spelling, punctuation, numbers, and line/paragraph breaks as faithfully as possible.
3. Do NOT translate, summarize, or add commentary. Output only the extracted text."""


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


def apply_doc_font(doc: Document, font_name: str = "Pyidaungsu", font_size: int = 11) -> None:
    """Apply font across default and common Word styles."""
    style_names = ["Normal", "Heading 1", "Heading 2", "Heading 3", "List Paragraph"]
    for name in style_names:
        if name in doc.styles:
            style = doc.styles[name]
            style.font.name = font_name
            if name == "Normal":
                style.font.size = Pt(font_size)
            rPr = style._element.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)
            rFonts.set(qn("w:ascii"), font_name)
            rFonts.set(qn("w:hAnsi"), font_name)
            rFonts.set(qn("w:cs"), font_name)


def render_page_png_bytes(page: fitz.Page, dpi: int = 250) -> bytes:
    """Render a PDF page to PNG bytes."""
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")


def list_available_gemini_models(api_key: str) -> List[str]:
    """List Gemini model names available for this API key."""
    client = genai.Client(api_key=api_key)
    models: List[str] = []
    for m in client.models.list():
        name = getattr(m, "name", "") or ""
        short_name = name.replace("models/", "", 1)
        if short_name and "gemini" in short_name.lower():
            models.append(short_name)

    # Prioritize flash/pro models if present
    preferred = [DEFAULT_MODEL]
    for candidate in models:
        if candidate not in preferred:
            preferred.append(candidate)
    return preferred


def _call_gemini_with_retry(
    client: genai.Client,
    png_bytes: bytes,
    model_name: str,
    max_retries: int = 2,
) -> str:
    """Call a specific Gemini model with short retries for transient errors."""
    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=png_bytes, mime_type="image/png"),
                    VISION_PROMPT,
                ],
            )
            return (response.text or "").strip()
        except APIError as exc:
            last_exc = exc
            code = getattr(exc, "code", None)
            if code in (503, 429, 500) and attempt < max_retries:
                time.sleep(attempt * 1.5)
                continue
            raise
    if last_exc:
        raise last_exc
    return ""


def gemini_vision_page(
    png_bytes: bytes,
    api_key: str,
    model_name: str = DEFAULT_MODEL,
) -> str:
    """Extract Myanmar + English text using Gemini Vision, trying available models if needed."""
    client = genai.Client(api_key=api_key)
    candidates: List[str] = [model_name]

    try:
        for available in list_available_gemini_models(api_key):
            if available not in candidates and ("flash" in available or "pro" in available):
                candidates.append(available)
    except Exception:
        pass

    last_exc: Optional[Exception] = None
    for candidate in candidates[:4]:
        try:
            return _call_gemini_with_retry(
                client=client,
                png_bytes=png_bytes,
                model_name=candidate,
                max_retries=2,
            )
        except APIError as exc:
            last_exc = exc
            code = getattr(exc, "code", None)
            if code in (404, 503, 429, 500):
                continue
            raise

    if last_exc:
        raise last_exc
    raise RuntimeError("No Gemini vision model succeeded for this page.")


def tesseract_ocr_page(png_bytes: bytes, lang: str = "mya+eng") -> str:
    """Fallback OCR using Tesseract."""
    img = Image.open(io.BytesIO(png_bytes))
    gray = ImageOps.grayscale(img)
    enhanced = ImageOps.autocontrast(gray)
    return pytesseract.image_to_string(enhanced, lang=lang, config="--psm 6").strip()


def extract_pages_text(
    pdf_bytes: bytes,
    use_gemini_vision: bool = True,
    gemini_api_key: Optional[str] = None,
    gemini_model: str = DEFAULT_MODEL,
    ocr_lang: str = "mya+eng",
    force_vision_or_ocr: bool = True,
    min_chars_threshold: int = 20,
    dpi: int = 250,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[str]:
    """Extract text per page using digital text first or Gemini Vision / Tesseract OCR."""
    api_key = (gemini_api_key or os.getenv("GEMINI_API_KEY") or "").strip()
    if use_gemini_vision and not api_key:
        raise ValueError("Gemini API key is missing. Enter it in the sidebar or set GEMINI_API_KEY.")

    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(pdf_doc)
    pages_text: List[str] = []

    try:
        for page_index in range(total_pages):
            page = pdf_doc[page_index]
            raw_text = "" if force_vision_or_ocr else page.get_text("text").strip()

            if force_vision_or_ocr or len(raw_text) < min_chars_threshold:
                png_bytes = render_page_png_bytes(page, dpi=dpi)
                if use_gemini_vision:
                    if progress_callback:
                        progress_callback(
                            page_index + 1,
                            total_pages,
                            f"Page {page_index + 1}/{total_pages}: Reading with Gemini Vision ({gemini_model})...",
                        )
                    text = gemini_vision_page(
                        png_bytes=png_bytes,
                        api_key=api_key,
                        model_name=gemini_model,
                    )
                else:
                    if progress_callback:
                        progress_callback(
                            page_index + 1,
                            total_pages,
                            f"Page {page_index + 1}/{total_pages}: Running Tesseract OCR ({ocr_lang})...",
                        )
                    text = tesseract_ocr_page(png_bytes=png_bytes, lang=ocr_lang)
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
    apply_doc_font(doc, font_name=font_name, font_size=font_size)

    total_pages = len(pages_text)
    for page_index, text in enumerate(pages_text):
        heading = doc.add_heading(f"Page {page_index + 1}", level=2)
        for run in heading.runs:
            set_run_font(run, font_name=font_name, font_size=font_size + 2)

        for line in text.splitlines():
            p = doc.add_paragraph()
            run = p.add_run(line)
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
    use_gemini_vision: bool = True,
    gemini_api_key: Optional[str] = None,
    gemini_model: str = DEFAULT_MODEL,
    ocr_lang: str = "mya+eng",
    force_vision_or_ocr: bool = True,
    dpi: int = 250,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> bytes:
    """Convert PDF bytes directly into Word (.docx) bytes."""
    pages_text = extract_pages_text(
        pdf_bytes=pdf_bytes,
        use_gemini_vision=use_gemini_vision,
        gemini_api_key=gemini_api_key,
        gemini_model=gemini_model,
        ocr_lang=ocr_lang,
        force_vision_or_ocr=force_vision_or_ocr,
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
    use_gemini_vision: bool = True,
    gemini_api_key: Optional[str] = None,
    gemini_model: str = DEFAULT_MODEL,
    ocr_lang: str = "mya+eng",
    force_vision_or_ocr: bool = True,
    dpi: int = 250,
) -> None:
    """Convert a PDF file on disk to a .docx file on disk."""
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_file}")

    docx_bytes = convert_pdf_bytes_to_docx_bytes(
        pdf_bytes=pdf_file.read_bytes(),
        font_name=font_name,
        font_size=font_size,
        use_gemini_vision=use_gemini_vision,
        gemini_api_key=gemini_api_key,
        gemini_model=gemini_model,
        ocr_lang=ocr_lang,
        force_vision_or_ocr=force_vision_or_ocr,
        dpi=dpi,
    )

    out_file = Path(docx_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_bytes(docx_bytes)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert Myanmar + English PDF files to Word (.docx) using Gemini Vision or OCR."
    )
    parser.add_argument("input_pdf", help="Path to input PDF file")
    parser.add_argument("output_docx", help="Path to output .docx file")
    parser.add_argument("--font", default="Pyidaungsu", help="Word font name (default: Pyidaungsu)")
    parser.add_argument("--size", type=int, default=11, help="Font size in pt (default: 11)")
    parser.add_argument("--api-key", default="", help="Gemini API key (or set GEMINI_API_KEY env var)")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Gemini model (default: {DEFAULT_MODEL})")
    parser.add_argument("--no-gemini", action="store_true", help="Use Tesseract OCR instead of Gemini Vision")
    parser.add_argument("--lang", default="mya+eng", help="Tesseract fallback language (default: mya+eng)")
    parser.add_argument("--dpi", type=int, default=250, help="Page render DPI (default: 250)")
    parser.add_argument("--digital-first", action="store_true", help="Use digital PDF text when available")

    args = parser.parse_args()
    convert_pdf_to_docx(
        pdf_path=args.input_pdf,
        docx_path=args.output_docx,
        font_name=args.font,
        font_size=args.size,
        use_gemini_vision=not args.no_gemini,
        gemini_api_key=args.api_key or None,
        gemini_model=args.model,
        ocr_lang=args.lang,
        force_vision_or_ocr=not args.digital_first,
        dpi=args.dpi,
    )
    print(f"Done! Saved Word document to: {args.output_docx}")
