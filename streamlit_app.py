import streamlit as st
import json
import re
from io import BytesIO
import fitz  # PyMuPDF
from google.cloud import vision
from google.oauth2 import service_account
import openai
import openpyxl


class CIMExtractor:
    """Main class for CIM financial data extraction."""
    
    def __init__(self):
        self.setup_credentials()
        self.setup_ui()
    
    def setup_credentials(self):
        """Initialize GCP and OpenAI credentials."""
        try:
            # GCP Credentials
            creds_dict = json.loads(st.secrets["GCP"]["gcp_credentials"])
            credentials = service_account.Credentials.from_service_account_info(creds_dict)
            self.vision_client = vision.ImageAnnotatorClient(credentials=credentials)
            
            # OpenAI credentials
            openai.api_key = st.secrets["OPENAI"]["OPENAI_API_KEY"]
            
        except Exception as e:
            st.error(f"❌ Error setting up credentials: {e}")
            st.stop()
    
    def setup_ui(self):
        """Setup Streamlit UI."""
        st.title("📊 CIM Financial Extractor (OCR + AI)")
        self.uploaded_pdf = st.file_uploader("📁 Upload CIM PDF", type=["pdf"])
    
    def join_wrapped_labels(self, text):
        """Join broken label lines like 'Adj. 4-Wall RR\nEBITDA' -> 'Adj. 4-Wall RR EBITDA'."""
        lines = text.split('\n')
        joined_lines = []
        buffer = ""

        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            if re.match(r"^[\d\$\(]", line):  # Starts with number/dollar — new data row
                if buffer:
                    joined_lines.append(buffer)
                    buffer = ""
                joined_lines.append(line)
            else:
                buffer = buffer + " " + line if buffer else line
        
        if buffer:
            joined_lines.append(buffer)

        return "\n".join(joined_lines)

    def preclean_combined_text(self, raw_text):
        """Clean and preprocess OCR text."""
        # Join broken label lines
        text = self.join_wrapped_labels(raw_text)
        
        # Add line breaks before each $number
        text = re.sub(r"(?<=\d)\s*(?=\$\d)", "\n", text)
        
        # Remove known junk headers
        junk_patterns = [
            r"\b(Joan Comp|OVERVIEW|onential - Not For Distribution|6\.|FINANCIAL)\b"
        ]
        for pattern in junk_patterns:
            text = re.sub(pattern, "", text)
        
        # Normalize extra spacing
        text = re.sub(r"\s{2,}", " ", text)
        
        return text

    def extract_text_from_pdf(self, pdf_bytes):
        """Extract text from PDF using OCR."""
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        combined_text = ""
        
        for i, page in enumerate(pdf_doc):
            st.text(f"Processing page {i+1} of {len(pdf_doc)}...")
            
            # Convert page to high-resolution image
            pix = page.get_pixmap(dpi=300)
            image_bytes = pix.tobytes("png")
            
            # OCR with Google Vision
            image = vision.Image(content=image_bytes)
            response = self.vision_client.document_text_detection(image=image)
            
            if response.error.message:
                st.warning(f"⚠️ OCR error on page {i+1}: {response.error.message}")
            else:
                combined_text += response.full_text_annotation.text + "\n"
        
        return self.preclean_combined_text(combined_text)

    def extract_financials_with_ai(self, text):
        """Extract financial data using GPT-4."""
        ai_prompt = self.build_ai_prompt(text)
        
        try:
            response = openai.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that ALWAYS responds with valid JSON."},
                    {"role": "user", "content": ai_prompt}
                ],
                temperature=0,
            )

            if not response or not response.choices or not response.choices[0].message.content:
                raise ValueError("Empty response from OpenAI API")

            response_text = response.choices[0].message.content.strip()
            
            if not response_text:
                raise ValueError("Empty response content from OpenAI API")

            return self.parse_ai_response(response_text)
            
        except Exception as e:
            st.error(f"❌ Error calling OpenAI API: {e}")
            st.stop()

    def build_ai_prompt(self, text):
        """Build the AI prompt for financial extraction."""
        return f"""
You are analyzing OCR output from a Confidential Information Memorandum (CIM) for an LBO model.

Extract the following hardcoded financials (not calculated, not inferred):

### Financial Metrics:
1. **Revenue** - Three most recent actual years + Six forward-looking years
2. **EBITDA** (prefer Adjusted or Run-Rate Adjusted if available) - Same structure as Revenue
3. **Maintenance CapEx** - Prefer values labeled "Maintenance CapEx" (not total CapEx)
4. **Acquisition Count** - Count of planned acquisitions per projected year

### Year Extraction Instructions:
- Identify all years tied to hardcoded financial values
- Sort chronologically and use:
  - Three earliest years for Actuals (Actual_1, Actual_2, Actual_3)
  - Next six years for forward-looking values (Expected, Proj_Y1 to Proj_Y5)

### Handling Year Labelss
- The first of the three Actual years should be written in Excel cell `E17` as the year
    - Example: if the three actuals are 2014, 2015, and 2016 → write 2014 in E17._

- Use the same year to generate the value for Excel cell `H17`:  
- Example: `FY2016A` → `LTM JUNE-16E` in H17.


### Candidate Metric Handling:
If multiple versions of a metric exist, group them in *_Candidates objects.

Example format:
```json
{{
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
    }}
  }}
}}
```

IMPORTANT: Respond with ONLY valid JSON. No explanatory text or markdown.
If no financial data found, return: {{"error": "No financial data found"}}

Text to analyze:
{text}
"""

    def parse_ai_response(self, response_text):
        """Parse and validate AI response."""
        st.subheader("📅 Raw GPT Response")
        st.code(response_text, language="text")

        # Clean JSON extraction
        cleaned_json_text = self.clean_json_response(response_text)
        
        if not cleaned_json_text:
            raise ValueError("No JSON content found in GPT response")

        st.subheader("📅 Extracted JSON")
        st.code(cleaned_json_text, language="json")

        try:
            data = json.loads(cleaned_json_text)
            
            if not isinstance(data, dict):
                raise ValueError(f"GPT response is not a valid JSON object. Got type: {type(data)}")
                
            if not data:
                st.warning("⚠️ GPT returned empty JSON object. No financial data found.")
                st.info("💡 Try uploading a different CIM document.")
                return {}
                
            return data
            
        except json.JSONDecodeError as e:
            st.error(f"❌ Failed to parse as JSON: {e}")
            st.error(f"Raw content: {repr(cleaned_json_text[:500])}")
            self.show_debug_info(response_text, cleaned_json_text)
            st.stop()

    def clean_json_response(self, response_text):
        """Clean JSON response from potential markdown formatting."""
        cleaned_text = response_text.strip()
        
        # Handle ```json blocks
        if "```json" in cleaned_text:
            match = re.search(r"```json(.*?)```", cleaned_text, re.DOTALL)
            if match:
                cleaned_text = match.group(1).strip()
        # Handle generic ``` blocks
        elif "```" in cleaned_text:
            match = re.search(r"```(.*?)```", cleaned_text, re.DOTALL)
            if match:
                cleaned_text = match.group(1).strip()
        
        return cleaned_text

    def show_debug_info(self, response_text, cleaned_text):
        """Show debug information for JSON parsing errors."""
        st.subheader("🔍 Debug Information")
        st.text("Full response text:")
        st.code(response_text, language="text")
        st.text("Cleaned JSON text:")
        st.code(cleaned_text, language="text")

    def flatten_financials(self, data):
        """Flatten nested metric dictionaries into one level."""
        flattened = {}
        
        for metric, values in data.items():
            if isinstance(values, dict):
                for subkey, value in values.items():
                    flattened[f"{metric}_{subkey}"] = value
            else:
                flattened[metric] = values
        
        return flattened

    def pick_metric_group(self, data, field_prefix, label):
        """Allow user to pick a single variant for metrics with multiple candidates."""
        candidates_key = f"{field_prefix}_Candidates"
        
        if candidates_key not in data:
            return
            
        st.subheader(f"🧐 Multiple variants found for {label}")
        choices = list(data[candidates_key].keys())
        
        if len(choices) == 1:
            selected = choices[0]
            st.info(f"✅ Only one {label} found: using '{selected}'")
        else:
            selected = st.radio(
                f"Choose one {label} version to use for ALL time periods:", 
                choices, 
                key=field_prefix
            )
        
        selected_values = data[candidates_key][selected]
        
        # Apply selected values to all time periods
        time_periods = [
            "Actual_1", "Actual_2", "Actual_3", "Expected",
            "Proj_Y1", "Proj_Y2", "Proj_Y3", "Proj_Y4", "Proj_Y5"
        ]
        
        for period in time_periods:
            field_name = f"{field_prefix}_{period}"
            if isinstance(selected_values, dict) and period in selected_values:
                data[field_name] = selected_values[period]

    def process_data(self, data):
        """Process and normalize extracted data."""
        # Flatten nested financials
        for field_prefix in ["Revenue", "Maintenance_CapEx", "Acquisition_Count"]:
            if field_prefix in data and isinstance(data[field_prefix], dict):
                for k, v in data[field_prefix].items():
                    data[f"{field_prefix}_{k}"] = v
        
        # Rename keys to match Excel mapping
        key_mappings = {
            "Maintenance_CapEx": "CapEx_Maint",
            "Acquisition_Count": "Num_Acq"
        }
        
        for old_prefix, new_prefix in key_mappings.items():
            suffixes = [
                "Actual_1", "Actual_2", "Actual_3", "Expected", 
                "Proj_Y1", "Proj_Y2", "Proj_Y3", "Proj_Y4", "Proj_Y5"
            ]
            for suffix in suffixes:
                old_key = f"{old_prefix}_{suffix}"
                new_key = f"{new_prefix}_{suffix}"
                if old_key in data:
                    data[new_key] = data[old_key]

        # Allow user to pick consistent metric types
        self.pick_metric_group(data, "EBITDA", "EBITDA")
        self.pick_metric_group(data, "Revenue", "Revenue")
        self.pick_metric_group(data, "CapEx_Maint", "Maintenance CapEx")
        self.pick_metric_group(data, "Num_Acq_Proj", "Acquisition Count")
        
        return data

    def get_excel_mapping(self):
        """Define Excel cell mapping for the LBO model."""
        return {
            # P&L Table (Historical)
            ("Revenue_Actual_1",): ("Model", "E20"),
            ("Revenue_Actual_2",): ("Model", "F20"),
            ("Revenue_Actual_3",): ("Model", "G20"),
            ("EBITDA_Actual_1",): ("Model", "E28"),
            ("EBITDA_Actual_2",): ("Model", "F28"),
            ("EBITDA_Actual_3",): ("Model", "G28"),
            
            # Projections (6 years forward)
            ("Revenue_Proj_Y1",): ("Model", "AC20"),
            ("Revenue_Proj_Y2",): ("Model", "AD20"),
            ("Revenue_Proj_Y3",): ("Model", "AE20"),
            ("Revenue_Proj_Y4",): ("Model", "AF20"),
            ("Revenue_Proj_Y5",): ("Model", "AG20"),
            ("Revenue_Proj_Y6",): ("Model", "AH20"),
            ("EBITDA_Proj_Y1",): ("Model", "AC28"),
            ("EBITDA_Proj_Y2",): ("Model", "AD28"),
            ("EBITDA_Proj_Y3",): ("Model", "AE28"),
            ("EBITDA_Proj_Y4",): ("Model", "AF28"),
            ("EBITDA_Proj_Y5",): ("Model", "AG28"),
            ("EBITDA_Proj_Y6",): ("Model", "AH28"),
            
            # Maintenance CapEx
            ("CapEx_Maint_Actual_1",): ("Model", "Z52"),
            ("CapEx_Maint_Actual_2",): ("Model", "AA52"),
            ("CapEx_Maint_Actual_3",): ("Model", "AB52"),
            ("CapEx_Maint_Proj_Y1",): ("Model", "AC52"),
            ("CapEx_Maint_Proj_Y2",): ("Model", "AD52"),
            ("CapEx_Maint_Proj_Y3",): ("Model", "AE52"),
            ("CapEx_Maint_Proj_Y4",): ("Model", "AF52"),
            ("CapEx_Maint_Proj_Y5",): ("Model", "AG52"),
            ("CapEx_Maint_Proj_Y6",): ("Model", "AH52"),
            
            # Acquisitions
            ("Num_Acq_Proj_Y1",): ("Acquisitions", "N13"),
            ("Num_Acq_Proj_Y2",): ("Acquisitions", "O13"),
            ("Num_Acq_Proj_Y3",): ("Acquisitions", "P13"),
            ("Num_Acq_Proj_Y4",): ("Acquisitions", "Q13"),
            ("Num_Acq_Proj_Y5",): ("Acquisitions", "R13"),
            
            # Headers
            ("Header_E17",): ("Model", "E17"),
            ("Header_H17",): ("Model", "H17"),
        }

    def update_excel_template(self, data):
        """Update Excel template with extracted data."""
        template_path = "TJC Practice Simple Model New (7).xlsx"
        
        try:
            wb = openpyxl.load_workbook(template_path)
            flattened_data = self.flatten_financials(data)
            mapping = self.get_excel_mapping()
            
            for key, (sheet_name, cell) in mapping.items():
                metric = key[0]
                if metric in flattened_data:
                    try:
                        wb[sheet_name][cell] = flattened_data[metric]
                    except Exception as e:
                        st.warning(f"⚠️ Failed to write {metric} → {sheet_name}!{cell}: {e}")
            
            # Save to BytesIO
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            return output
            
        except FileNotFoundError:
            st.error(f"❌ Excel template not found: {template_path}")
            return None
        except Exception as e:
            st.error(f"❌ Error updating Excel template: {e}")
            return None

    def run(self):
        """Main execution flow."""
        if not self.uploaded_pdf:
            return
            
        # Extract text from PDF
        with st.spinner("🧠 OCRing the CIM..."):
            pdf_bytes = self.uploaded_pdf.read()
            combined_text = self.extract_text_from_pdf(pdf_bytes)
        
        st.success("✅ OCR complete!")
        
        # Show OCR output
        st.subheader("🔍 Full OCR Text")
        with st.expander("Click to view OCR output"):
            st.text(combined_text)
        
        # Extract financials with AI
        with st.spinner("🔍 Extracting financial metrics with GPT-4..."):
            data = self.extract_financials_with_ai(combined_text)
        
        if not data:
            return
            
        # Process and normalize data
        processed_data = self.process_data(data)
        
        # Update Excel template
        excel_output = self.update_excel_template(processed_data)
        
        if excel_output:
            st.download_button(
                label="📅 Download Updated LBO Excel",
                data=excel_output,
                file_name="updated_lbo_model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# Initialize and run the application
if __name__ == "__main__":
    extractor = CIMExtractor()
    extractor.run()
