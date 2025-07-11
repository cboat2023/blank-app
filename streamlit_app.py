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
import re

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

def pick_metric_group(field_prefix, label):
    """
    If *_Candidates exists (e.g., EBITDA_Candidates), ask user to pick a single variant,
    then apply that value to all matching keys like EBITDA_Actual_1, _2, Expected, Proj_Y1–5.
    """
    candidates_key = f"{field_prefix}_Candidates"
    if candidates_key in data:
        st.subheader(f"🧐 Multiple variants found for {label}")
        choices = list(data[candidates_key].keys())
        selected = st.radio(f"Choose one {label} version to use for ALL time periods:", choices, key=field_prefix)
        selected_values = data[candidates_key][selected]

        for subfield in [
            "Actual_1", "Actual_2", "Expected",
            "Proj_Y1", "Proj_Y2", "Proj_Y3", "Proj_Y4", "Proj_Y5"
        ]:
            field_name = f"{field_prefix}_{subfield}"
            if isinstance(selected_values, dict) and field_name in selected_values:
                data[field_name] = selected_values[field_name]

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

    with st.spinner("🔍 Extracting financial metrics with GPT-4..."):
        ai_prompt = f"""
You are analyzing OCR output from a Confidential Information Memorandum (CIM) for an LBO model.

Your task is to extract the following **hardcoded** financials (not calculated, not inferred):

### Financial Metrics:
1. **Revenue**
   - Two most recent actual years (e.g., 2023A, 2024A)
   - One expected/budget year (e.g., 2025E)
   - Five projected years (e.g., 2026E to 2030E)

2. **EBITDA** (prefer Adjusted or RR Adj.)
   - Same format: 2 recent actuals, 1 expected, 5 projected

3. **Maintenance CapEx**
   - Prefer labeled “Maintenance CapEx” (not total CapEx)
   - Same format: 2 actual, 1 expected, 5 projected

4. **Acquisition Count**
   - Count of planned acquisitions per projected year
   - If none are explicitly listed, say \"assumed\": 1 for each year

### Candidate Handling Instructions:
If **multiple values** are found for any metric (e.g., both "Adj. EBITDA" and "Reported EBITDA"), return all values in a `*_Candidates` field, like this:

```json
"EBITDA_Candidates": {{
  "Adj. EBITDA": {{"EBITDA_Actual_1": ..., "EBITDA_Proj_Y5": ...}},
  "Reported EBITDA": {{"EBITDA_Actual_1": ..., "EBITDA_Proj_Y5": ...}}
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
        st.subheader("📅 Extracted Financial Metrics (JSON)")
        st.code(response_text, language="json")

        cleaned_json_text = response_text
        if "```json" in cleaned_json_text:
            match = re.search(r"```json(.*?)```", cleaned_json_text, re.DOTALL)
            if match:
                cleaned_json_text = match.group(1).strip()
        else:
            cleaned_json_text = cleaned_json_text.strip()

        try:
            data = json.loads(cleaned_json_text)
        except Exception as e:
            st.error(f"❌ Failed to parse GPT response as JSON:\n{e}")
            st.stop()

        # Allow user to pick consistent metric types
        pick_metric_group("EBITDA", "EBITDA")
        pick_metric_group("Revenue", "Revenue")
        pick_metric_group("CapEx_Maint", "Maintenance CapEx")
        pick_metric_group("Num_Acq_Proj", "Acquisition Count")

        # Excel cell mapping
        mapping = {
            ("Revenue_Actual_1",): ("Model", "E20"),
            ("Revenue_Actual_2",): ("Model", "F20"),
            ("Revenue_Expected",): ("Model", "G20"),

            ("EBITDA_Actual_1",): ("Model", "E28"),
            ("EBITDA_Actual_2",): ("Model", "F28"),
            ("EBITDA_Expected",): ("Model", "G28"),

            ("Revenue_Proj_Y1",): ("Model", "AC20"),
            ("Revenue_Proj_Y2",): ("Model", "AD20"),
            ("Revenue_Proj_Y3",): ("Model", "AE20"),
            ("Revenue_Proj_Y4",): ("Model", "AF20"),
            ("Revenue_Proj_Y5",): ("Model", "AG20"),

            ("EBITDA_Proj_Y1",): ("Model", "AC28"),
            ("EBITDA_Proj_Y2",): ("Model", "AD28"),
            ("EBITDA_Proj_Y3",): ("Model", "AE28"),
            ("EBITDA_Proj_Y4",): ("Model", "AF28"),
            ("EBITDA_Proj_Y5",): ("Model", "AG28"),

            ("CapEx_Maint_Actual_1",): ("Model", "AA52"),
            ("CapEx_Maint_Actual_2",): ("Model", "AB52"),
            ("CapEx_Maint_Expected",): ("Model", "AC52"),
            ("CapEx_Maint_Proj_Y1",): ("Model", "AD52"),
            ("CapEx_Maint_Proj_Y2",): ("Model", "AE52"),
            ("CapEx_Maint_Proj_Y3",): ("Model", "AF52"),
            ("CapEx_Maint_Proj_Y4",): ("Model", "AG52"),
            ("CapEx_Maint_Proj_Y5",): ("Model", "AH52"),

            ("Num_Acq_Proj_Y1",): ("Acquisitions", "N13"),
            ("Num_Acq_Proj_Y2",): ("Acquisitions", "O13"),
            ("Num_Acq_Proj_Y3",): ("Acquisitions", "P13"),
            ("Num_Acq_Proj_Y4",): ("Acquisitions", "Q13"),
            ("Num_Acq_Proj_Y5",): ("Acquisitions", "R13"),
        }

        template_path = "TJC Practice Simple Model New (7).xlsx"
        wb = openpyxl.load_workbook(template_path)

        for key, (sheet_name, cell) in mapping.items():
            metric = key[0]
            if metric in data:
                try:
                    wb[sheet_name][cell] = data[metric]
                except Exception as e:
                    st.warning(f"⚠️ Failed to write {metric} → {sheet_name}!{cell}: {e}")

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📅 Download Updated LBO Excel",
            data=output,
            file_name="updated_lbo_model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


