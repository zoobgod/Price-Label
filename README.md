# Pharma PI + Label Generator (Streamlit)

This app ingests:
- PI/Invoice PDF
- MSDS PDF
- Specification PDF

Then it extracts customs-relevant fields and generates:
1. `Proforma_Invoice.docx`
2. `Label.docx`

## What it extracts
- From PI: product rows (positions), quantity, unit/total price, currency, buyer/exporter, invoice no/date
- From Specification: terms of delivery, period of validity, specification date
- From MSDS: storage temperature (OCR-enabled for scanned PDFs)

## OCR approach
- Native PDF text extraction first (`pypdf`)
- OCR fallback (`PyMuPDF` rendering + `pytesseract`) for low-text/scanned pages
- PI/Invoice OCR can be force-enabled from UI (enabled by default)

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Install Tesseract engine (required for scanned PDFs):

```bash
# macOS
brew install tesseract

# Optional extra language pack
brew install tesseract-lang
```

## Run

```bash
streamlit run app.py
```

## Template support (optional)
You can upload custom `.docx` templates for Price List and Label.

Use placeholders like:
- `{{INVOICE_NO}}`
- `{{INVOICE_DATE}}`
- `{{BUYER_NAME}}`
- `{{BUYER_ADDRESS}}`
- `{{EXPORTER_NAME}}`
- `{{EXPORTER_ADDRESS}}`
- `{{TERMS_OF_DELIVERY}}`
- `{{PERIOD_OF_VALIDITY}}`
- `{{SPECIFICATION_DATE}}`
- `{{STORAGE_TEMPERATURE}}`
- `{{COMPANY_INFO}}`
- `{{POSITIONS_TABLE}}`
- `{{POSITION_1_NAME_EN}}`
- `{{POSITION_1_NAME_RU}}`
- `{{POSITION_1_QUANTITY}}`
- `{{POSITION_1_PACKING_EN}}`
- `{{POSITION_1_PACKING_RU}}`
- `{{POSITION_1_PRICE}}`
- `{{POSITION_1_TOTAL}}`
- `{{POSITION_1_CURRENCY}}`

Template behavior:
- First pass: explicit `{{KEY}}` placeholders
- Second pass: semantic key matching (e.g., `PRODUCT NAME:`, `QUANTITY:`, `PACKING:`, `TERMS OF DELIVERY:`, `Storage:`)
- If no fields can be mapped, the app automatically falls back to built-in generated documents (it will not return an unchanged template).

## Notes
- Extraction quality depends on source PDF quality and consistent formatting.
- Always review fields in the UI before generating final customs documents.
