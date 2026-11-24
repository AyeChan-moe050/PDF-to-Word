from pathlib import Path
import streamlit as st
from converter import convert_pdf_bytes_to_docx_bytes

st.set_page_config(
    page_title="Myanmar + English PDF to Word",
    page_icon="📄",
    layout="centered",
)

st.title("Myanmar + English PDF to Word Converter")
st.write("Upload a PDF containing Myanmar and/or English text to convert it into an editable `.docx` file.")

with st.sidebar:
    st.header("Conversion Settings")
    font_name = st.text_input("Word Font Name", value="Pyidaungsu")
    font_size = st.slider("Font Size (pt)", min_value=9, max_value=24, value=11)
    ocr_lang = st.selectbox(
        "OCR Languages",
        options=["mya+eng", "mya", "eng"],
        index=0,
    )
    force_ocr = st.checkbox(
        "Force OCR on all pages",
        value=False,
        help="Enable for scanned PDFs or PDFs with broken/legacy font encoding.",
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
            with st.spinner("Converting PDF..."):
                docx_bytes = convert_pdf_bytes_to_docx_bytes(
                    pdf_bytes=uploaded_file.getvalue(),
                    font_name=font_name.strip() or "Pyidaungsu",
                    font_size=font_size,
                    ocr_lang=ocr_lang,
                    force_ocr=force_ocr,
                    progress_callback=on_progress,
                )

            output_filename = f"{Path(uploaded_file.name).stem}.docx"
            st.success("Conversion complete!")
            st.download_button(
                label="Download Word Document (.docx)",
                data=docx_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as exc:
            st.error(f"Conversion failed: {exc}")
