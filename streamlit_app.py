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

def join_wrapped_labels(text):
    """
    Joins broken label lines like:
    'Adj. 4-Wall RR\nEBITDA' -> 'Adj. 4-Wall RR EBITDA'
    """
    # Join lines where a line ends without a number and the next line starts with a capital letter
    lines = text.split('\n')
    joined_lines = []
    buffer = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if re.match(r"^[\d\$\(]", line):  # starts with a number/dollar — new data row
            if buffer:
                joined_lines.append(buffer)
                buffer = ""
            joined_lines.append(line)
        else:
            if buffer:
                buffer += " " + line
            else:
                buffer = line
    if buffer:
        joined_lines.append(buffer)

    return "\n".join(joined_lines)


def preclean_combined_text(raw_text):
    # Step 1: Join broken label lines
    text = join_wrapped_labels(raw_text)

    # Step 2: Add line breaks before each $number
    text = re.sub(r"(?<=\d)\s*(?=\$\d)", "\n", text)

    # Step 3: Remove known junk headers
    text = re.sub(r"\b(Joan Comp|OVERVIEW|onential - Not For Distribution|6\.|FINANCIAL)\b", "", text)

    # Step 4: Normalize extra spacing
    text = re.sub(r"\s{2,}", " ", text)

    return text
    
def flatten_financials(data):
    """Flatten nested metric dictionaries (like EBITDA, Revenue) into one level."""
    flattened = {}

    for metric, values in data.items():
        if isinstance(values, dict):
            for subkey, value in values.items():
                flattened[f"{metric}_{subkey}"] = value
        else:
            flattened[metric] = values  # Handle flat keys like Fiscal_Years if needed

    return flattened


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
        if len(choices) == 1:
            selected = choices[0]
            st.info(f"✅ Only one {label} found: using “{selected}”")
        selected_values = data[candidates_key][selected]

        for subfield in [
            "Actual_1", "Actual_2", "Expected",
            "Proj_Y1", "Proj_Y2", "Proj_Y3", "Proj_Y4", "Proj_Y5"
        ]:
            field_name = f"{field_prefix}_{subfield}"
            if isinstance(selected_values, dict) and subfield in selected_values:
                data[field_name] = selected_values[subfield]


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

    combined_text = preclean_combined_text(combined_text)


    # After combined_text has been built
    st.subheader("🔍 Full OCR Text")
    with st.expander("Click to view OCR output"):
        st.text(combined_text)


    with st.spinner("🔍 Extracting financial metrics with GPT-4..."):
        ai_prompt = f"""
You are analyzing OCR output from a Confidential Information Memorandum (CIM) for an LBO model.

Your task is to extract the following **hardcoded** financials (not calculated, not inferred):

---

### Financial Metrics to Extract:

1. **Revenue**
   - Three most recent actual years ( ex: FY2022A, 2022A, 2022 Actual)
   - 6 projected/ expected years (e.g., 2025E to 2030E or FY2025E to FY2030E )

2. **EBITDA** (prefer Adjusted or Run-Rate Adjusted)
   - Same format: 3 recent actuals,6 expected/projected

3. **Maintenance CapEx**
   - Prefer values explicitly labeled “Maintenance CapEx” (do not infer from Total CapEx)
   - Same format: 3 recent actuals,6 expected/projected

4. **Acquisition Count**
   - Count of planned acquisitions per projected year (only if explicitly stated)
   - If not found, assume 1 acquisition per projected year and return: "assumed"

---

### Special Fields for Excel Auto-Population:

5. **Historical Year Header for Excel (E17):**
   - Extract the **first of three most recent historical year** you will find the historicals written down as formatted ex: FY2021A,FY2022A, FY2023A choose FY2021A (e.g., 2022 if years are 2022, 2023, 2024)
   - Return  in this format: '2022', just input the year

6. **LTM Label (H17):**
   - Take the **next year after the most recent actual** (e.g., 2025 if last actual is 2024)
   - Return as string in this format: `"LTM JUNE-25E"` under the key `"Header_H17"`

---

### Candidate Handling Instructions:

If multiple types of a metric are found (e.g., "Adj. EBITDA" and "Reported EBITDA"), provide them inside a `*_Candidates` field. Each entry should be a dictionary with values for all 9 periods:

- `Actual_1`, `Actual_2`, `Actual_3`
- `Expected`
- `Proj_Y1`, ..., `Proj_Y5`

For example:

```json
"EBITDA_Candidates": {{
  "Adj. EBITDA": {{
    "Actual_1": 25.1,
    "Actual_2": 27.3,
    "Actual_3": 29.1,
    "Expected": 30.0,
    "Proj_Y1": 32.5,
    "Proj_Y2": 35.0,
    "Proj_Y3": 37.5,
    "Proj_Y4": 40.0,
    "Proj_Y5": 42.0
  }},
  "Reported EBITDA": {{
    "Actual_1": 22.4,
    "Actual_2": 23.9,
    "Actual_3": 24.5,
    "Expected": 26.0,
    "Proj_Y1": 28.0,
    "Proj_Y2": 29.5,
    "Proj_Y3": 31.0,
    "Proj_Y4": 32.5,
    "Proj_Y5": 34.0
  }}
}}


IMPORTANT: You MUST respond with ONLY valid JSON. Do not include any explanatory text, markdown formatting, or code blocks. Start directly with the opening curly brace and end with the closing curly brace.

If no financial data is found, return: {{"error": "No financial data found"}}

Text to analyze:
{combined_text}
"""

        try:
            response = openai.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that ALWAYS responds with valid JSON."},
                    {"role": "user", "content": ai_prompt}
                ],
                temperature=0,
            )

            # Check if response exists and has content
            if not response or not response.choices or not response.choices[0].message.content:
                st.error("❌ Empty response from OpenAI API")
                st.stop()

            response_text = response.choices[0].message.content.strip()
            
            # Check if response is empty
            if not response_text:
                st.error("❌ Empty response content from OpenAI API")
                st.stop()

            st.subheader("📅 Raw GPT Response")
            st.code(response_text, language="text")

            # Clean JSON extraction
            cleaned_json_text = response_text
            if "```json" in cleaned_json_text:
                match = re.search(r"```json(.*?)```", cleaned_json_text, re.DOTALL)
                if match:
                    cleaned_json_text = match.group(1).strip()
                else:
                    st.error("❌ Found ```json marker but couldn't extract JSON content")
                    st.stop()
            elif "```" in cleaned_json_text:
                # Handle case where there's ``` but no json marker
                match = re.search(r"```(.*?)```", cleaned_json_text, re.DOTALL)
                if match:
                    cleaned_json_text = match.group(1).strip()
            else:
                cleaned_json_text = cleaned_json_text.strip()

            # Additional check for empty cleaned text
            if not cleaned_json_text:
                st.error("❌ No JSON content found in GPT response")
                st.stop()

            st.subheader("📅 Extracted JSON")
            st.code(cleaned_json_text, language="json")

            # Parse JSON with better error handling
            try:
                data = json.loads(cleaned_json_text)
                
                # Validate that we got a dictionary
                if not isinstance(data, dict):
                    st.error("❌ GPT response is not a valid JSON object")
                    st.error(f"Got type: {type(data)}")
                    st.stop()
                    
                # Check if we got any meaningful data
                if not data:
                    st.warning("⚠️ GPT returned empty JSON object. This might indicate no financial data was found in the document.")
                    st.info("💡 Try uploading a different CIM document or check if the document contains clear financial tables.")
                    data = {}  # Initialize empty dict to prevent errors below
                    
            except json.JSONDecodeError as e:
                st.error(f"❌ Failed to parse as JSON: {e}")
                st.error(f"Raw content to parse: {repr(cleaned_json_text[:500])}")
                
                # Show the user what we actually received
                st.subheader("🔍 Debug Information")
                st.text("Full response text:")
                st.code(response_text, language="text")
                st.text("Cleaned JSON text:")
                st.code(cleaned_json_text, language="text")
                st.stop()
                
            except Exception as e:
                st.error(f"❌ Unexpected error parsing JSON: {e}")
                st.stop()

        except Exception as e:
            st.error(f"❌ Error calling OpenAI API: {e}")
            st.stop()

        # Flatten nested financials
        for field_prefix in ["Revenue", "Maintenance_CapEx", "Acquisition_Count"]:
            if field_prefix in data and isinstance(data[field_prefix], dict):
                for k, v in data[field_prefix].items():
                    data[f"{field_prefix}_{k}"] = v
        
        # Rename keys to match Excel mapping
        for old_prefix, new_prefix in [
            ("Maintenance_CapEx", "CapEx_Maint"),
            ("Acquisition_Count", "Num_Acq")
        ]:
            for suffix in ["Actual_1", "Actual_2", "Actual_3", "Expected", "Proj_Y1", "Proj_Y2", "Proj_Y3", "Proj_Y4", "Proj_Y5"]:
                old_key = f"{old_prefix}_{suffix}"
                new_key = f"{new_prefix}_{suffix}"
                if old_key in data:
                    data[new_key] = data[old_key]


        # Allow user to pick consistent metric types
        pick_metric_group("EBITDA", "EBITDA")
        pick_metric_group("Revenue", "Revenue")
        pick_metric_group("CapEx_Maint", "Maintenance CapEx")
        pick_metric_group("Num_Acq_Proj", "Acquisition Count")

        # Excel cell mapping
        mapping = {
    # P&L Table (Historical): E-G columns
    ("Revenue_Actual_1",): ("Model", "E20"),    # Oldest actual (e.g., 2022A)
    ("Revenue_Actual_2",): ("Model", "F20"),    # Middle actual (e.g., 2023A)
    ("Revenue_Actual_3",): ("Model", "G20"),    # Most recent actual (e.g., 2024A)

    ("EBITDA_Actual_1",): ("Model", "E28"),
    ("EBITDA_Actual_2",): ("Model", "F28"),
    ("EBITDA_Actual_3",): ("Model", "G28"),

    # Management Projection Table: 1 expected + 5 projections
    ("Revenue_Expected",): ("Model", "AC20"),   # Expected year (e.g., 2025E)
    ("Revenue_Proj_Y1",): ("Model", "AD20"),
    ("Revenue_Proj_Y2",): ("Model", "AE20"),
    ("Revenue_Proj_Y3",): ("Model", "AF20"),
    ("Revenue_Proj_Y4",): ("Model", "AG20"),
    ("Revenue_Proj_Y5",): ("Model", "AH20"),

    ("EBITDA_Expected",): ("Model", "AC28"),
    ("EBITDA_Proj_Y1",): ("Model", "AD28"),
    ("EBITDA_Proj_Y2",): ("Model", "AE28"),
    ("EBITDA_Proj_Y3",): ("Model", "AF28"),
    ("EBITDA_Proj_Y4",): ("Model", "AG28"),
    ("EBITDA_Proj_Y5",): ("Model", "AH28"),

    # Maintenance CapEx in projection table
    ("CapEx_Maint_Actual_2",): ("Model", "AA52"),   # Skipping Actual_1 due to missing mapping
    ("CapEx_Maint_Actual_3",): ("Model", "AB52"),
    ("CapEx_Maint_Expected",): ("Model", "AC52"),
    ("CapEx_Maint_Proj_Y1",): ("Model", "AD52"),
    ("CapEx_Maint_Proj_Y2",): ("Model", "AE52"),
    ("CapEx_Maint_Proj_Y3",): ("Model", "AF52"),
    ("CapEx_Maint_Proj_Y4",): ("Model", "AG52"),
    ("CapEx_Maint_Proj_Y5",): ("Model", "AH52"),

    # Acquisition counts (projections only)
    ("Num_Acq_Proj_Y1",): ("Acquisitions", "N13"),
    ("Num_Acq_Proj_Y2",): ("Acquisitions", "O13"),
    ("Num_Acq_Proj_Y3",): ("Acquisitions", "P13"),
    ("Num_Acq_Proj_Y4",): ("Acquisitions", "Q13"),
    ("Num_Acq_Proj_Y5",): ("Acquisitions", "R13"),


    ("EBITDA_Actual_3",): ("Model", "G28"),   # Most recent actual EBITDA
    ("Header_E17",): ("Model", "E17"),        # Third most recent historical year
    ("Header_H17",): ("Model", "H17"),        # LTM header


}



        flattened_data = flatten_financials(extracted_data)

        template_path = "TJC Practice Simple Model New (7).xlsx"
        wb = openpyxl.load_workbook(template_path)

        for key, (sheet_name, cell) in mapping.items():
            metric = key[0]
            if metric in flattened_data:
                try:
                    wb[sheet_name][cell] = flattened_data[metric]
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




