import os
import tempfile
from pdf2docx import Converter
import streamlit as st

st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

st.title("📄 PDF to Word Converter")
st.write(
    "Upload a PDF file below to convert it into an editable Word document (.docx)."
)

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Selected file: **{uploaded_file.name}**")

    if st.button("Convert to Word (.docx)", type="primary"):
        with st.spinner("Converting document... please wait"):
            # Save uploaded PDF to a temporary file
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".pdf"
            ) as temp_in:
                temp_in.write(uploaded_file.read())
                temp_in_path = temp_in.name

            # Target output docx file path
            temp_out_path = temp_in_path.replace(".pdf", ".docx")

            try:
                # Perform the conversion
                cv = Converter(temp_in_path)
                cv.convert(temp_out_path)
                cv.close()

                # Read converted file to provide download
                with open(temp_out_path, "rb") as f:
                    docx_bytes = f.read()

                st.success("Conversion completed successfully!")
                st.download_button(
                    label="⬇️ Download Word Document",
                    data=docx_bytes,
                    file_name=uploaded_file.name.replace(".pdf", ".docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"Error during conversion: {e}")
            finally:
                # Cleanup temporary files
                if os.path.exists(temp_in_path):
                    os.remove(temp_in_path)
                if os.path.exists(temp_out_path):
                    os.remove(temp_out_path)
