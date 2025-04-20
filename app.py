import streamlit as st
from pdf2docx import Converter
import pypandoc
import os

st.set_page_config(page_title="Word ↔ PDF Converter", layout="centered")

st.title("📄 Word ↔ PDF Converter")
st.markdown("Upload your file and convert between DOCX and PDF formats.")

# Upload file
uploaded_file = st.file_uploader("Choose a DOCX or PDF file", type=["docx", "pdf"])

# Choose conversion direction
conversion_type = st.radio("Convert to:", ("PDF (from DOCX)", "DOCX (from PDF)"))

if uploaded_file:
    filename = uploaded_file.name
    with open(filename, "wb") as f:
        f.write(uploaded_file.read())

    # Convert button
    if st.button("🔄 Convert"):
        if conversion_type == "PDF (from DOCX)" and filename.endswith(".docx"):
            try:
                output_file = filename.replace(".docx", ".pdf")
                pypandoc.convert_file(filename, "pdf", outputfile=output_file)
                with open(output_file, "rb") as f:
                    st.success("Conversion successful!")
                    st.download_button("⬇️ Download PDF", f, file_name=output_file)
            except Exception as e:
                st.error(f"Conversion failed: {e}")

        elif conversion_type == "DOCX (from PDF)" and filename.endswith(".pdf"):
            try:
                output_file = filename.replace(".pdf", ".docx")
                cv = Converter(filename)
                cv.convert(output_file, start=0, end=None)
                cv.close()
                with open(output_file, "rb") as f:
                    st.success("Conversion successful!")
                    st.download_button("⬇️ Download DOCX", f, file_name=output_file)
            except Exception as e:
                st.error(f"Conversion failed: {e}")
        else:
            st.warning("Please upload the correct file type for the selected conversion.")

    # Clean up files (optional)
    if os.path.exists(filename):
        os.remove(filename)
