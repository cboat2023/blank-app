import streamlit as st
import json
import io
import fitz  # PyMuPDF
from google.cloud import vision
from google.oauth2 import service_account
from openai import OpenAI

# --- GCP Credentials from secrets ---
creds_dict = json.loads(st.secrets["GCP"]["gcp_credentials"])
credentials = service_account.Credentials.from_service_account_info(creds_dict)
client = vision.ImageAnnotatorClient(credentials=credentials)

# --- OpenAI credentials ---
openai_api_key = st.secrets["OPENAI_API_KEY"]
openai_client = OpenAI(api_key=openai_api_key)

# --- Streamlit UI ---
st.title("📊 CIM Financial Extractor (OCR + AI)")
uploaded_pdf = st.file_uploader("📁 Upload CIM PDF", type=["pdf"])

if uploaded_pdf:
    with st.spinner("🧠 OCRing the CIM..."):
        # Read and convert PDF
        pdf_bytes = uploaded_pdf.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        combined_text = ""
        for i, page in enumerate(pdf_doc):
            st.text(f"OCRing page {i+1} of {len(pdf_doc)}...")
            pix = page.get_pixmap(dpi=300)
            image_bytes = pix.tobytes("png")

            image = vision.Image(content=image_bytes)
            response = client.document_text_detection(image=image)

            if response.error.message:
                st.warning(f"⚠️ OCR error on page {i+1}: {response.error.message}")
            else:
                combined_text += response.full_text_annotation.text + "\n"

    st.success("✅ OCR complete!")

    # --- AI Financial Extraction ---
    with st.spinner("🔍 Extracting financial metrics with GPT-4..."):

        ai_prompt = f"""
You are analyzing OCR output from a Confidential Information Memorandum (CIM) for an LBO model.

Your job is to extract the following **hardcoded** financials (not calculated, not inferred):

1. Revenue
   - Two actual years
   - One expected/budget year
   - Five projected years
   - Also: any new region / acquisition revenue

2. EBITDA (prefer Adjusted or RR Adj.)
   - Same format: 2 actual, 1 expected, 5 forecast

3. Maintenance CapEx
   - Prefer labeled “Maintenance CapEx”, not total CapEx
   - 2 actual, 1 expected, 5 projected

4. Acquisition Count per projected year

Return your answer in valid JSON using this structure:
```json
{{
  "Revenue_Actual_1": ..., "Revenue_Actual_2": ..., "Revenue_Expected": ..., 
  "Revenue_Proj_Y1": ..., "Revenue_Proj_Y2": ..., "Revenue_Proj_Y3": ..., 
  "Revenue_Proj_Y4": ..., "Revenue_Proj_Y5": ...,

  "Revenue_Acq_Actual_1": ..., "Revenue_Acq_Actual_2": ..., 
  "Revenue_Acq_Expected": ..., "Revenue_Acq_Proj_Y1": ..., 
  "Revenue_Acq_Proj_Y2": ..., "Revenue_Acq_Proj_Y3": ..., 
  "Revenue_Acq_Proj_Y4": ..., "Revenue_Acq_Proj_Y5": ...,

  "EBITDA_Actual_1": ..., "EBITDA_Actual_2": ..., "EBITDA_Expected": ..., 
  "EBITDA_Proj_Y1": ..., "EBITDA_Proj_Y2": ..., "EBITDA_Proj_Y3": ..., 
  "EBITDA_Proj_Y4": ..., "EBITDA_Proj_Y5": ...,

  "CapEx_Maint_Actual_1": ..., "CapEx_Maint_Actual_2": ..., 
  "CapEx_Maint_Expected": ..., "CapEx_Maint_Proj_Y1": ..., 
  "CapEx_Maint_Proj_Y2": ..., "CapEx_Maint_Proj_Y3": ..., 
  "CapEx_Maint_Proj_Y4": ..., "CapEx_Maint_Proj_Y5": ...,

  "Num_Acq_Proj_Y1": ..., "Num_Acq_Proj_Y2": ..., "Num_Acq_Proj_Y3": ..., 
  "Num_Acq_Proj_Y4": ..., "Num_Acq_Proj_Y5": ...
}}
Text to analyze:
{combined_text}
"""

        response = openai_client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": ai_prompt}
            ],
            temperature=0,
        )

        response_text = response.choices[0].message.content.strip()

        st.subheader("📥 Extracted Financial Metrics (JSON)")
        st.code(response_text, language="json")
