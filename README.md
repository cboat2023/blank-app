# 📊 CIM-to-LBO Automation App

This Streamlit-based application automates the process of extracting key financial metrics from Confidential Information Memorandums (CIMs) and mapping them into a pre-formatted Leveraged Buyout (LBO) Excel model.

## 🚀 Features

- **Smart File Handling**:
  - 📄 Digital PDFs → parsed using `pdfplumber`
  - 📄 Scanned PDFs → processed with Google Vision OCR
  - 🖼️ Image uploads (JPG/PNG) → OCR extraction

- **AI-Powered Financial Extraction**:
  - Uses GPT-4 to extract key financials (Revenue, Adjusted EBITDA, Maintenance CapEx, Acquisition Counts)
  - Handles multiple metric candidates and lets the user choose which to use

- **Excel Integration**:
  - Automatically maps extracted data into a custom LBO model Excel template
  - Downloads a completed Excel file with one click

## 📁 Supported Inputs

- Digital CIM PDFs (text-based)
- Scanned CIM PDFs (image-based)
- Images of financial tables (JPG/PNG)

## 📦 Technologies Used

- `Streamlit` – UI and deployment
- `pdfplumber` – Digital PDF text and table extraction
- `PyMuPDF` – Scanned PDF structure and page detection
- `Google Cloud Vision` – OCR for scanned PDFs and images
- `OpenAI GPT-4` – Financial data extraction
- `openpyxl` – Excel automation

## 🧠 How It Works

1. Upload a PDF or image
2. App detects file type and extracts raw text
3. GPT-4 parses financials into structured JSON
4. You confirm metric variants (if needed)
5. Outputs a fully populated Excel LBO model

## 📂 Excel Template

This app uses a pre-defined LBO Excel template named:
``TJC Practice Simple Model New (7).xlsx``
Make sure this file exists in the root directory before launching the app.

## 🔧 Setup Instructions

1. Clone the repo
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
