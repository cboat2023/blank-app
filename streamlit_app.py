from google.cloud import vision
from google.oauth2 import service_account
import streamlit as st
import json
import io
from pdf2image import convert_from_bytes
from PIL import Image

# --- Load credentials from Streamlit secrets ---
creds_dict = json.loads(st.secrets["GCP"]["gcp_credentials"])
credentials = service_account.Credentials.from_service_account_info(creds_dict)

# --- Initialize Vision client with credentials ---
client = vision.ImageAnnotatorClient(credentials=credentials)

# --- Streamlit UI ---
st.title("CIM OCR Extractor")
uploaded_pdf = st.file_uploader("Upload CIM PDF", type=["pdf"])

if uploaded_pdf:
    with st.spinner("Converting PDF to images..."):
        # Save the uploaded PDF temporarily
        pdf_path = "temp_uploaded.pdf"
        with open(pdf_path, "wb") as f:
            f.write(uploaded_pdf.read())

        from pdf2image import convert_from_bytes

        # Read uploaded file directly
        pdf_bytes = uploaded_pdf.read()

        

        # Convert PDF to image pages
        pages = convert_from_bytes(uploaded_pdf.read(), dpi=300)


        # Initialize Vision API client
        client = vision.ImageAnnotatorClient()

        combined_text = ""
        for i, page in enumerate(pages):
            st.text(f"OCRing page {i+1} of {len(pages)}")
            img_byte_arr = io.BytesIO()
            page.save(img_byte_arr, format='PNG')
            image = vision.Image(content=img_byte_arr.getvalue())
            response = client.document_text_detection(image=image)

            if response.error.message:
                st.warning(f"Error on page {i+1}: {response.error.message}")
            else:
                combined_text += response.full_text_annotation.text + "\n"

        # Save OCR output
        output_path = "ocr_combined.txt"
        with open(output_path, "w") as f:
            f.write(combined_text)

        st.success("✅ OCR complete.")
        st.download_button("Download OCR Text", combined_text, file_name="ocr_combined.txt", mime="text/plain")
