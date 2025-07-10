import streamlit as st
import json
import io
import fitz  # PyMuPDF
from google.cloud import vision
from google.oauth2 import service_account

# --- Load credentials from Streamlit secrets ---
creds_dict = json.loads(st.secrets["GCP"]["gcp_credentials"])
credentials = service_account.Credentials.from_service_account_info(creds_dict)

# --- Initialize Vision client with credentials ---
client = vision.ImageAnnotatorClient(credentials=credentials)

# --- Streamlit UI ---
st.title("📄 CIM OCR Extractor")
uploaded_pdf = st.file_uploader("Upload CIM PDF", type=["pdf"])

if uploaded_pdf:
    with st.spinner("🔄 Converting PDF pages to images and running OCR..."):
        # Load PDF in memory with fitz (PyMuPDF)
        pdf_bytes = uploaded_pdf.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        combined_text = ""

        for i, page in enumerate(pdf_doc):
            st.text(f"🧠 OCRing page {i+1} of {len(pdf_doc)}")
            pix = page.get_pixmap(dpi=300)
            image_bytes = pix.tobytes("png")

            image = vision.Image(content=image_bytes)
            response = client.document_text_detection(image=image)

            if response.error.message:
                st.warning(f"⚠️ OCR error on page {i+1}: {response.error.message}")
            else:
                combined_text += response.full_text_annotation.text + "\n"

        st.success("✅ OCR complete!")

        # --- Offer download ---
        st.download_button(
            label="⬇️ Download OCR Text",
            data=combined_text,
            file_name="ocr_combined.txt",
            mime="text/plain"
        )

