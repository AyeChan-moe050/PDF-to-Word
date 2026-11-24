import os
from pathlib import Path
import streamlit as st
from converter import DEFAULT_MODEL, convert_pdf_bytes_to_docx_bytes

st.set_page_config(
    page_title="Myanmar + English PDF to Word",
    page_icon="📄",
    layout="centered",
)

st.title("Myanmar + English PDF to Word Converter")
st.write("Upload a PDF file and convert it into a `.docx` file with **Pyidaungsu** font.")

env_api_key = os.getenv("GEMINI_API_KEY", "")

with st.sidebar:
    st.header("Recognition Engine")
    engine = st.radio(
        "Choose Engine",
        options=["Gemini Vision (Recommended)", "Tesseract OCR"],
        index=0,
    )
    use_gemini = engine.startswith("Gemini Vision")

    gemini_api_key = ""
    gemini_model = DEFAULT_MODEL
    ocr_lang = "mya+eng"

    if use_gemini:
        gemini_api_key = st.text_input(
            "Gemini API Key",
            value=env_api_key,
            type="password",
            help="Enter your Google Gemini API key (or set GEMINI_API_KEY in environment / Streamlit Secrets).",
        )
        gemini_model = st.selectbox(
            "Gemini Model",
            options=["gemini-3.8-flash", "gemini-3.8-pro"],
            index=0,
        )
    else:
        ocr_lang = st.selectbox(
            "OCR Languages",
            options=["mya+eng", "mya", "eng"],
            index=0,
        )

    st.header("Word Formatting")
    font_name = st.text_input("Word Font Name", value="Pyidaungsu")
    font_size = st.slider("Font Size (pt)", min_value=9, max_value=24, value=11)
    force_vision_or_ocr = st.checkbox(
        "Process all pages with Vision/OCR",
        value=True,
        help="Recommended when PDF text extraction has broken Myanmar font encoding.",
    )

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    if st.button("Convert to Word (.docx)", type="primary"):
        progress_bar = st.progress(0)
        status_text = st.empty()

        def on_progress(current: int, total: int, message: str) -> None:
            progress_bar.progress(current / total)
            status_text.text(message)

        try:
            with st.spinner("Converting PDF to Word..."):
                docx_bytes = convert_pdf_bytes_to_docx_bytes(
                    pdf_bytes=uploaded_file.getvalue(),
                    font_name=font_name.strip() or "Pyidaungsu",
                    font_size=font_size,
                    use_gemini_vision=use_gemini,
                    gemini_api_key=gemini_api_key.strip() or None,
                    gemini_model=gemini_model,
                    ocr_lang=ocr_lang,
                    force_vision_or_ocr=force_vision_or_ocr,
                    progress_callback=on_progress,
                )

            output_filename = f"{Path(uploaded_file.name).stem}.docx"
            st.success("Done! Download your Word file below.")
            st.download_button(
                label="Download Word Document (.docx)",
                data=docx_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
        except Exception as exc:
            st.error(f"Conversion failed: {exc}")
