import streamlit as st
import io
import os
from google.cloud import vision
from pdf2image import convert_from_bytes

# --- Authenticate with GCP ---
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "analog-height-465315-p7-3b93a3344a21.json"  # Upload this file to Streamlit Cloud separately

# --- Streamlit UI ---
st.title("📊 CIM OCR to Financials Extractor")

uploaded_files = st.file_uploader("Upload screenshots or PDF (up to 5)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)

if uploaded_files:
    client = vision.ImageAnnotatorClient()
    full_text = ""

    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        st.write(f"📄 Processing: {filename}")

        if filename.endswith(".pdf"):
            pages = convert_from_bytes(uploaded_file.read())
            for i, page in enumerate(pages):
                buf = io.BytesIO()
                page.save(buf, format="PNG")
                image = vision.Image(content=buf.getvalue())
                response = client.document_text_detection(image=image)
                full_text += response.full_text_annotation.text + "\n"
        else:
            image = vision.Image(content=uploaded_file.read())
            response = client.document_text_detection(image=image)
            full_text += response.full_text_annotation.text + "\n"

    st.text_area("📝 Extracted OCR Text", full_text, height=300)
