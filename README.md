# Pharma PI + Label Generator (Streamlit)

This app ingests:
- PI/Invoice PDF
- MSDS PDF
- Specification PDF

It generates:
1. `Price_List.docx`
2. Label output(s):
- `Label.docx` when all positions share one storage temperature
- `Label_<temperature>.docx` per temperature group when positions differ

## Extraction behavior
- PI extraction runs native text + OCR candidate and auto-selects better structured parse.
- MSDS/spec extraction supports OCR for scanned PDFs.
- If no storage temperature is found, default is `+15C to +25C ambient`.

## Install (local)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Install OCR engine:

```bash
# macOS
brew install tesseract
brew install tesseract-lang
```

Run:

```bash
streamlit run app.py
```

## Streamlit Cloud
This repo includes `packages.txt` for system dependencies:
- `tesseract-ocr`
- `tesseract-ocr-rus`

Streamlit Cloud installs these automatically on deploy.

## Template placeholders
The app supports fixed placeholders used by your templates, including:
- `{{INVOICE_DATE}}`
- `{{TERMS_OF_DELIVERY}}`
- `{{PERIOD_OF_VALIDITY}}`
- `{{STORAGE_TEMPERATURE_EN}}`
- `{{STORAGE_TEMPERATURE_RU}}`
- `{{EXPORTER_COMPANY_NAME_EN}}`
- `{{EXPORTER_COMPANY_NAME_RU}}`
- `{{EXPORTER_COMPANY_ADRESS_EN}}`
- `{{POSITION_1_NAME_EN}} ... {{POSITION_N_NAME_EN}}`
- `{{POSITION_1_NAME_RU}} ... {{POSITION_N_NAME_RU}}`
- `{{POSITION_1_QUANTITY_EN}} ... {{POSITION_N_QUANTITY_EN}}`
- `{{POSITION_1_QUANTITY_RU}} ... {{POSITION_N_QUANTITY_RU}}`
- `{{POSITION_1_PACKING_EN}} ... {{POSITION_N_PACKING_EN}}`
- `{{POSITION_1_PACKING_RU}} ... {{POSITION_N_PACKING_RU}}`
- `{{POSITION_1_UNIT_PRICE}} ... {{POSITION_N_UNIT_PRICE}}`
- `{{POSITION_1_TOTAL_PRICE}} ... {{POSITION_N_TOTAL_PRICE}}`

Unresolved placeholders are cleared automatically.

## Notes
- Use the Positions section to set `storage_temperature` per position if products must be split by temperature.
- Always review extracted fields before final export.
- Workflow is step-based: Extract -> Review -> Generate.
- Generated download buttons remain visible until `Start From Beginning` is clicked.
