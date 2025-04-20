import streamlit as st
from pdf2docx import Converter
from docx2pdf import convert
import os
import dropbox

# === CONFIG ===
dropbox_token = os.environ.get("sl.u.AFo_SWgXgbhGHXYUAXO2bjHzsh9vdhzVIxXMNvvYZhFuaJeB99EXfBbaHfyN_Nhshm09pNkx4GUPaEDtDNU5AGsLmAmWpawp6uT7F2Nhp1wkJtSjpllPIS5dpJc4T-GMJqb8qKR5n-rKJMnoDyCnnMaZdNjPLyWn0wMII7-xDxqvHN7Ahn7bXNMuT3pfI0A2mqmmV3l5dqf1uLnee-PWPrjNZ7xEZqCqp0MG-G1zMJTEE13wKl6gRpizFA3CyJjNVHQOHBITE3GJzkID3PRSTZTH7sF8mIX-bY0W_hON2vVKAH0kUUUex697uRRvohkQQY5xES9AU9COWeoHLmL-zdySNjjy4RCNlFfoLGyMMUQkU4CgVwxarB-MhPIo90Pxhwg8mDiW50RcGJHeIak_5PcGfKmgOT19WsfNGblR8UeVOcMIUU_8Cy44JApoFlIacqKJoUQNuepPN8NN7WdswnJyJjqlJCWYNNAuDePe4FSjlGSsyTNsJmz_BpcFdxNdJbbuyHggm_J-k2zG-x-QQIFHe-V2jIQn7igaYMOamua7BeFjV0nGT0mU83TkHEh-hxAIiNx0AuU-snZftuDBGfFXwBWo-mgUrshVdu2OTxVefEwp0c2qAs9rWN3zOV0NjaX4YgNdgR0mVnlj-3a1TTccDYHiMsyYBcSwx7WzyMVhWVaq2SIMiYumQy_xE2j8Gw0OP6bIXnmM0KJxKnquwXzyDQbtNwcp9AlCPoTmkO5F3LvDMDi6-C8U60KyZs4g7C4aVVF7-2z1Azrwxw7YYUjax0GN-26ZGha8luR1j9wbfZvEOuGIWN33P2o5asmjxFeBo3PF2Qq-S_S-DELuSj0Ly2SMhZo-NsICN8WYQTCVpt0a6vqXwGuS74oY_5XPshqvk8z8S1z4I7O9-L8R-gfA3m4ezFaD-BwkveNK6Hg_eLS7MZtmps7tn_67uO81GTN4_nXLf06FFG6fRdNZ2ob_OM8Ul7L1hQKFrXRQY7RYdAubhSDrVXO2kfOqfIPCvzDVnGmeIJ5pHY0R9Mub9X_TmK3C20a2VZbUJm2ssKV0TEo9O5aic3AjZrtUp4n8LSb_bc5N9k94-2XS9XwSb-FCRoOlzmth5vWzqJSuFF1J--Lz_q44aOulMG9bdF0euGfx0Wf_RY5NAyUdk0877q7B5bIVYSYbikV3xXppmqh6XYzptemWxg59CR3MXRGaMT8a4XBVQQwdxMP0QPtNLbMz60eFGKxKAKvqFsTEHnGs3PupPJMUnTB67nsYjVb4b2l-Li_BU9XtgHtXiDdPT1P6npvi41NX3Yjz9p7F7pI1YTOcR6Ybr0OUbmOFvtEy7chCbF5euZSODSe6EfkWEkT9E1qR6vkeUeGmEqu_OqpZU9hHhFX15V6ytzYvM5GSBVFIDxAhDwEWFv6M4LmWGZZA8bwY67QZ6xp-r34YptmneA")  # Ensure this is set in your environment
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
