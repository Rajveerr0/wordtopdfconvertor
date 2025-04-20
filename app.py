import streamlit as st
from pdf2docx import Converter
import pypandoc
import os
import dropbox

# === CONFIG ===
DROPBOX_ACCESS_TOKEN = os.getenv("sl.u.AFrYaE5UWF9QUvUfmc6FIj7AY1-Mjlg1xXCHYzhNPCyQEn5_Av5ewd8honxAdaNTpQ5l5wvYxVJTxFUPqeep8CnV3Cyml8v234NviQnF_M5JJGh7hGs_CjHmg9yTNzcKHH0QEQogL6RdNOqS8UY1zxPhBNAzbAv6Q6xTF-6lBwlaRIZX_QCLZ5_NQihYnc-FlfZTCsU7CFNzs_GSLSkyc5mEFqi_eQ0USbgm0fDYKWl6CMi0W7MpzL00JE7tUihF4hk1xrGb5N-H4CCfu9X6jvN3YBcVDx8l8NUDpTJM7mCqfLBFdjHEDsCt8bPyGrgHsnvVUzQnDO7EpSR5T3HyPDt9UfY8HXmeFJYC96UhLcR-LKdpMqPp_Ae-bZOUOgTbxDC6W5p_L0I-PTsmcMSEdsz4LpAFI_b0lBhWby9jh69FK2w2iRtaYSZQZlygClwr29oTWYNsxdLSAsgAAca6cUsvoX-0eZ3TRd0T9XqZ5o1NAAVeGQUSWohB-3120HkwxB5RQGoBKUebZEWDegznA3sf_WsecJ0CvrpEzROUeYj9Y_vs3gDTdQR-jKHTEVCnoIoqcBTkBKcyd8kH8xEF4e2fuRqcnf-qJbMQvFojb2Xi3wU9LByQsSBYySeD_pI1opfB-f1PCoXhtXpa1sSudo6znwh_P5s70EXKZd5bCnRFr8_vbKHvnLBzxumiXPWS0fi90TmauU_Rq-LkSC64mZIFqPw3WhSN7AjC_nlJ-rygm7xK9MqqsyB1RLIjlNX-vBnMwk0Dl_jpsj1Hr2HY_7UIlQtyRuIuS8X51gL-CCLVeR3Ip6WUwo6Zt4wZvM8CdQkHWXfJYis_jIHPlijtIYaejkV2IG7ngR7Ki8O7aLscylNx-QHMPYRIZdqS2RkATC3PZg59aPn7E5aNQqUc3aabWmniqU8tzKRZ9CVVaVr8zCG3JAjbP1O9hVvF3K3i8TYW8UR2nlU1SeMctut0LcBuZELxWy3XDk8YjRyQk8ANAHP1w1eKdtd4ydniGgivX0Tzym5epxHLWidL_vLWxHoEBauD1-uO2qvHjwTLuuXmav0bsZOKn8lHieXT20fI7xEhnkttQlGDjHpAzNxi5rwSpqqSmQjmWHRv32saPTVNJ1c1N4g159NEXB6-7X5Z4g3jG711U5T8qdOKkLvi-5aTdqfxC24IuopcMfCd5gSZiXL96v5xVWLUEX2dr05Od1HG-xpIyoJE5QqkdsSavQIOxltk6x16ZdoRKRTFZ_LRp3uOU1tClZUlMRKweCQgAvRcdbkJCSJurSL3wqi17D0R833PkvRDSuXJ2Ak6JGFR67pChm2r7xmXvwuPXYp8KoIKBGDGks2dGCas3Eafi9jHNd52BMW25i9_SFclvIUhDB6GxNXZ4V52MIBh_3rBwfTfLXHBMXaYbkkYENKUPrwDQDDaroQl4-RVyhR3w3XOOw")

st.set_page_config(page_title="Word ↔ PDF Converter", layout="centered")
st.title("📄 Word ↔ PDF Converter")
st.markdown("Upload your file and convert between DOCX and PDF formats.")

# === Dropbox Upload Function ===
def upload_to_dropbox(local_path, dropbox_path):
    try:
        dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
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
                pypandoc.convert_file(filename, "pdf", outputfile=output_file)
                with open(output_file, "rb") as f:
                    st.success("Conversion successful!")
                    st.download_button("⬇️ Download PDF", f, file_name=output_file)
                # Upload to Dropbox
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
                # Upload to Dropbox
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
