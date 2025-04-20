import streamlit as st
from pdf2docx import Converter
from docx2pdf import convert
import os
import dropbox

# === CONFIG ===
dropbox_token = os.environ.get("DROPBOX_ACCESS_TOKEN")  # Ensure this is set in your environment
dbx = dropbox.Dropbox(dropbox_token)

st.set_page_config(page_title="Word ↔ PDF Converter", layout="centered")
st.title("📄 Word ↔ PDF Converter")
st.markdown("Upload your file and convert between DOCX and PDF formats.")

# === Dropbox Upload Function ===
def upload_to_dropbox(local_path, dropbox_path):
    try:
        with open(local_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)
        shared_link = dbx.sharing_create_shared_link_with_settings(dropbox_path)
        return shared_link.url
    except Exception as e:
        return f"Error uploading to Dropbox: {e}"

# === File Upload UI ===
uploaded_file = st.file_uploader("Choose a DOCX or PDF file", type=["docx", "pdf"])
conversion_type = st.radio("Convert to:", ("PDF (from DOCX)", "DOCX (from PDF)"))

if uploaded_file:
    filename = uploaded_file.name
    with open(filename, "wb") as f:
        f.write(uploaded_file.read())

    if st.button("🔄 Convert"):
        if conversion_type == "PDF (from DOCX)" and filename.endswith(".docx"):
            try:
                output_file = filename.replace(".docx", ".pdf")
                convert(filename, output_file)
                with open(output_file, "rb") as f:
                    st.success("Conversion successful!")
                    st.download_button("⬇️ Download PDF", f, file_name=output_file)
                # Dropbox
                dropbox_link = upload_to_dropbox(output_file, f"/{output_file}")
                st.markdown(f"[📥 Dropbox Link]({dropbox_link})")
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
                # Dropbox
                dropbox_link = upload_to_dropbox(output_file, f"/{output_file}")
                st.markdown(f"[📥 Dropbox Link]({dropbox_link})")
            except Exception as e:
                st.error(f"Conversion failed: {e}")
        else:
            st.warning("Please upload the correct file type for the selected conversion.")

    # Clean up files
    if os.path.exists(filename):
        os.remove(filename)
    if 'output_file' in locals() and os.path.exists(output_file):
        os.remove(output_file)
