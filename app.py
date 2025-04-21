import streamlit as st
from pdf2docx import Converter
import mammoth
import weasyprint
import os
import dropbox

# === CONFIG ===
dropbox_token = os.environ.get("DROPBOX_ACCESS_TOKEN")  # Set this in Render's environment variables
dbx = dropbox.Dropbox(dropbox_token)

st.set_page_config(page_title="Word ↔ PDF Converter", layout="centered")
st.title("📄 Word ↔ PDF Converter")
st.markdown("Upload your file and convert between DOCX and PDF formats.")

# === Dropbox Upload Function ===
def upload_to_dropbox(local_path, dropbox_path):
    try:
        with open(local_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)

        # Try to create a new shared link
        try:
            shared_link = dbx.sharing_create_shared_link_with_settings(dropbox_path)
        except dropbox.exceptions.ApiError as e:
            # If shared link already exists, fetch it
            if isinstance(e.error, dropbox.sharing.CreateSharedLinkWithSettingsError) and e.error.is_shared_link_already_exists():
                links = dbx.sharing_list_shared_links(path=dropbox_path, direct_only=True).links
                shared_link = links[0] if links else None
            else:
                raise e

        return shared_link.url if shared_link else "No shared link available."
    except Exception as e:
        return f"Error uploading to Dropbox: {e}"

# === DOCX to PDF Conversion (Linux-Compatible) ===
def convert_docx_to_pdf(input_path, output_path):
    with open(input_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value
        with open("temp.html", "w", encoding="utf-8") as html_file:
            html_file.write(html)
        weasyprint.HTML("temp.html").write_pdf(output_path)

# === File Upload UI ===
uploaded_file = st.file_uploader("Choose a DOCX or PDF file", type=["docx", "pdf"])
conversion_type = st.radio("Convert to:", ("PDF (from DOCX)", "DOCX (from PDF)"))

if uploaded_file:
    filename = uploaded_file.name
    with open(filename, "wb") as f:
        f.write(uploaded_file.read())

    if st.button("🔄 Convert"):
        output_file = ""
        try:
            if conversion_type == "PDF (from DOCX)" and filename.endswith(".docx"):
                output_file = filename.replace(".docx", ".pdf")
                convert_docx_to_pdf(filename, output_file)
                st.success("Conversion successful!")
                with open(output_file, "rb") as f:
                    st.download_button("⬇️ Download PDF", f, file_name=output_file)

            elif conversion_type == "DOCX (from PDF)" and filename.endswith(".pdf"):
                output_file = filename.replace(".pdf", ".docx")
                cv = Converter(filename)
                cv.convert(output_file, start=0, end=None)
                cv.close()
                st.success("Conversion successful!")
                with open(output_file, "rb") as f:
                    st.download_button("⬇️ Download DOCX", f, file_name=output_file)

            else:
                st.warning("Please upload the correct file type for the selected conversion.")
                output_file = ""

            # Upload to Dropbox if conversion was successful
            if output_file and os.path.exists(output_file):
                dropbox_link = upload_to_dropbox(output_file, f"/{output_file}")
                st.markdown(f"[📥 Dropbox Link]({dropbox_link})")

        except Exception as e:
            st.error(f"Conversion failed: {e}")

        # Clean up temp files
        if os.path.exists(filename):
            os.remove(filename)
        if output_file and os.path.exists(output_file):
            os.remove(output_file)
        if os.path.exists("temp.html"):
            os.remove("temp.html")
