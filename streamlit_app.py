import streamlit as st
import json
import io
import fitz  # PyMuPDF
from google.cloud import vision
from google.oauth2 import service_account
import openai
import pandas as pd
import openpyxl
from io import BytesIO

# --- GCP Credentials from secrets ---
creds_dict = json.loads(st.secrets["GCP"]["gcp_credentials"])
credentials = service_account.Credentials.from_service_account_info(creds_dict)
client = vision.ImageAnnotatorClient(credentials=credentials)

# --- OpenAI credentials ---
openai_api_key = st.secrets["OPENAI"]["OPENAI_API_KEY"]
openai.api_key = openai_api_key

# --- Streamlit UI ---
st.title("📊 CIM Financial Extractor (OCR + AI)")
uploaded_pdf = st.file_uploader("📁 Upload CIM PDF", type=["pdf"])

if uploaded_pdf:
    with st.spinner("🧠 OCRing the CIM..."):
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

1. Revenue – 2 actual, 1 expected, 5 projected
2. EBITDA – same format
3. Maintenance CapEx – 2 actual, 1 expected, 5 projected
4. Acquisition Count – 5 projected years (if missing, use 1 per year)

Return ONLY a valid JSON object with `null` as default for missing values.
DO NOT include markdown, backticks, or explanations.

Example format:
{{
  "Revenue_Actual_1": null,
  "Revenue_Actual_2": null,
  "Revenue_Expected": null,
  "Revenue_Proj_Y1": null,
  ...
}}

Text to analyze:
{combined_text}
"""

        response = openai.chat.completions.create(
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

        # Try to parse GPT output
        try:
            data = json.loads(response_text)
        except Exception as e:
            st.error("❌ Failed to parse GPT response as JSON")
            st.stop()

        # === Excel mapping dictionary ===
        mapping = {
    # Revenue (Actual + Expected)
    ("Revenue_Actual_1",): ("Model", "E20"),
    ("Revenue_Actual_2",): ("Model", "F20"),
    ("Revenue_Expected",): ("Model", "G20"),

    # EBITDA (Actual + Expected)
    ("EBITDA_Actual_1",): ("Model", "E28"),
    ("EBITDA_Actual_2",): ("Model", "F28"),
    ("EBITDA_Expected",): ("Model", "G28"),

    # Revenue (Projected)
    ("Revenue_Proj_Y1",): ("Model", "AC20"),
    ("Revenue_Proj_Y2",): ("Model", "AD20"),
    ("Revenue_Proj_Y3",): ("Model", "AE20"),
    ("Revenue_Proj_Y4",): ("Model", "AF20"),
    ("Revenue_Proj_Y5",): ("Model", "AG20"),

    # EBITDA (Projected)
    ("EBITDA_Proj_Y1",): ("Model", "AC28"),
    ("EBITDA_Proj_Y2",): ("Model", "AD28"),
    ("EBITDA_Proj_Y3",): ("Model", "AE28"),
    ("EBITDA_Proj_Y4",): ("Model", "AF28"),
    ("EBITDA_Proj_Y5",): ("Model", "AG28"),

    # Maintenance CapEx
    ("CapEx_Maint_Actual_1",): ("Model", "AA52"),
    ("CapEx_Maint_Actual_2",): ("Model", "AB52"),
    ("CapEx_Maint_Expected",): ("Model", "AC52"),
    ("CapEx_Maint_Proj_Y1",): ("Model", "AD52"),
    ("CapEx_Maint_Proj_Y2",): ("Model", "AE52"),
    ("CapEx_Maint_Proj_Y3",): ("Model", "AF52"),
    ("CapEx_Maint_Proj_Y4",): ("Model", "AG52"),
    ("CapEx_Maint_Proj_Y5",): ("Model", "AH52"),

    # Acquisition Count (Projected) — Acquisitions sheet
    ("Num_Acq_Proj_Y1",): ("Acquisitions", "N13"),
    ("Num_Acq_Proj_Y2",): ("Acquisitions", "O13"),
    ("Num_Acq_Proj_Y3",): ("Acquisitions", "P13"),
    ("Num_Acq_Proj_Y4",): ("Acquisitions", "Q13"),
    ("Num_Acq_Proj_Y5",): ("Acquisitions", "R13"),
        }

        # === Load Excel Template ===
        template_path = "TJC Practice Simple Model New (7).xlsx"
        wb = openpyxl.load_workbook(template_path)

        for key, (sheet_name, cell) in mapping.items():
            metric = key[0]
            if metric in data:
                value = data[metric]
                try:
                    wb[sheet_name][cell] = value
                except Exception as e:
                    st.warning(f"Failed to write {metric} → {sheet_name}!{cell}: {e}")

        # Save & download updated Excel
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download Updated LBO Excel",
            data=output,
            file_name="updated_lbo_model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
