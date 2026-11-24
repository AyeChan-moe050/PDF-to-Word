from pathlib import Path
import streamlit as st
from converter import extract_pages_text, pages_text_to_docx_bytes

st.set_page_config(
    page_title="Myanmar + English PDF to Word",
    page_icon="📄",
    layout="wide",
)

st.title("Myanmar + English PDF to Word Converter")
st.write(
    "Preview extracted/OCR text first, edit if needed, and download the `.docx` only when it looks right."
)

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
        value=True,
        help="Recommended for scanned PDFs or PDFs where Myanmar text does not extract properly.",
    )

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if "pages_text" not in st.session_state:
    st.session_state.pages_text = None
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "last_file_key" not in st.session_state:
    st.session_state.last_file_key = None

if uploaded_file is not None:
    file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.last_file_key != file_key:
        st.session_state.last_file_key = file_key
        st.session_state.pages_text = None
        st.session_state.docx_bytes = None

    if st.button("1. Extract & Preview Text", type="primary"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        st.session_state.docx_bytes = None

        def on_progress(current: int, total: int, message: str) -> None:
            progress_bar.progress(current / total)
            status_text.text(message)

        try:
            with st.spinner("Reading PDF..."):
                st.session_state.pages_text = extract_pages_text(
                    pdf_bytes=uploaded_file.getvalue(),
                    ocr_lang=ocr_lang,
                    force_ocr=force_ocr,
                    progress_callback=on_progress,
                )
            st.success("Preview ready! Check the text below before generating Word.")
        except Exception as exc:
            st.error(f"Failed to read PDF: {exc}")

    if st.session_state.pages_text is not None:
        st.subheader("Text Preview (Editable)")
        edited_pages = []
        for idx, page_content in enumerate(st.session_state.pages_text):
            edited = st.text_area(
                label=f"Page {idx + 1}",
                value=page_content,
                height=220,
                key=f"page_editor_{file_key}_{idx}",
            )
            edited_pages.append(edited)

        if st.button("2. Confirm Text & Generate Word (.docx)"):
            try:
                st.session_state.docx_bytes = pages_text_to_docx_bytes(
                    pages_text=edited_pages,
                    font_name=font_name.strip() or "Pyidaungsu",
                    font_size=font_size,
                )
                st.success("Word file is ready! Click Download below.")
            except Exception as exc:
                st.error(f"Failed to build Word file: {exc}")

    if st.session_state.docx_bytes is not None:
        output_filename = f"{Path(uploaded_file.name).stem}.docx"
        st.download_button(
            label="3. Download Word Document (.docx)",
            data=st.session_state.docx_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
