from __future__ import annotations

from dataclasses import asdict
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import streamlit as st

from src.doc_builder import generate_label_doc, generate_price_list_doc
from src.models import ExtractedData, Position
from src.pipeline import run_extraction_pipeline

st.set_page_config(page_title="Pharma PI + Label Generator", layout="wide")


def _positions_to_df(positions: list[Position]) -> pd.DataFrame:
    records = [asdict(p) for p in positions] if positions else [asdict(Position())]
    return pd.DataFrame(records)


def _df_to_positions(df: pd.DataFrame) -> list[Position]:
    rows = []
    for _, row in df.fillna("").iterrows():
        quantity = row.get("quantity", "")
        unit_price = row.get("unit_price", "")
        total_price = row.get("total_price", "")

        def _to_float(v):
            if v in ("", None):
                return None
            try:
                return float(str(v).replace(",", ""))
            except ValueError:
                return None

        rows.append(
            Position(
                code=str(row.get("code", "")).strip(),
                name_en=str(row.get("name_en", "")).strip(),
                name_ru=str(row.get("name_ru", "")).strip(),
                quantity=_to_float(quantity),
                packing_en=str(row.get("packing_en", "")).strip(),
                packing_ru=str(row.get("packing_ru", "")).strip(),
                unit_price=_to_float(unit_price),
                total_price=_to_float(total_price),
                currency=str(row.get("currency", "")).strip(),
            )
        )
    return rows


def _pack_outputs(price_docx: bytes, label_docx: bytes) -> bytes:
    bio = BytesIO()
    with ZipFile(bio, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("Proforma_Invoice.docx", price_docx)
        zf.writestr("Label.docx", label_docx)
    return bio.getvalue()


st.title("Pharmacopeia Customs Document Builder")
st.caption("Upload PI + MSDS + Specification PDFs, review extracted data, then generate Price List and Label files.")

with st.sidebar:
    st.header("Inputs")
    pi_pdf = st.file_uploader("PI / Invoice PDF", type=["pdf"], key="pi_pdf")
    msds_pdf = st.file_uploader("MSDS PDF", type=["pdf"], key="msds_pdf")
    spec_pdf = st.file_uploader("Specification PDF", type=["pdf"], key="spec_pdf")

    st.header("Optional Templates")
    st.caption(
        "Template placeholders supported: {{INVOICE_NO}}, {{INVOICE_DATE}}, {{BUYER_NAME}}, {{TERMS_OF_DELIVERY}},"
        " {{PERIOD_OF_VALIDITY}}, {{SPECIFICATION_DATE}}, {{STORAGE_TEMPERATURE}}, {{COMPANY_INFO}},"
        " {{POSITIONS_TABLE}}, {{POSITION_1_NAME_EN}}, {{POSITION_1_NAME_RU}}, {{POSITION_1_QUANTITY}}"
    )
    price_template = st.file_uploader("Price List .docx template", type=["docx"], key="price_tpl")
    label_template = st.file_uploader("Label .docx template", type=["docx"], key="label_tpl")

    st.header("OCR")
    force_msds_ocr = st.checkbox("Force OCR on MSDS", value=True)

    run = st.button("Extract Documents", type="primary", use_container_width=True)

if run:
    if not pi_pdf:
        st.error("PI/Invoice PDF is required.")
    else:
        with st.spinner("Extracting fields from PDFs..."):
            extracted, logs = run_extraction_pipeline(
                pi_pdf_bytes=pi_pdf.getvalue(),
                msds_pdf_bytes=msds_pdf.getvalue() if msds_pdf else None,
                specification_pdf_bytes=spec_pdf.getvalue() if spec_pdf else None,
                force_ocr_msds=force_msds_ocr,
            )
        st.session_state["extracted"] = extracted
        st.session_state["logs"] = logs
        msds_meta = (logs.get("msds") or {}).get("meta") or {}
        if msds_pdf and not msds_meta.get("tesseract_available", False):
            st.warning(
                "Tesseract OCR is not available on this machine. Scanned MSDS files may not extract storage temperature."
            )

if "extracted" in st.session_state:
    extracted: ExtractedData = st.session_state["extracted"]

    tab_review, tab_positions, tab_generate, tab_debug = st.tabs(
        ["Review Fields", "Positions", "Generate", "Extraction Log"]
    )

    with tab_review:
        c1, c2 = st.columns(2)
        with c1:
            extracted.invoice_no = st.text_input("Invoice No", value=extracted.invoice_no)
            extracted.invoice_date = st.text_input("Invoice Date", value=extracted.invoice_date)
            extracted.buyer_name = st.text_input("Buyer Name", value=extracted.buyer_name)
            extracted.buyer_address = st.text_area("Buyer Address", value=extracted.buyer_address, height=120)
            extracted.exporter_name = st.text_input("Exporter Name", value=extracted.exporter_name)
            extracted.exporter_address = st.text_area("Exporter Address", value=extracted.exporter_address, height=120)

        with c2:
            extracted.terms_of_delivery = st.text_input("Terms Of Delivery", value=extracted.terms_of_delivery)
            extracted.period_of_validity = st.text_input("Period Of Validity", value=extracted.period_of_validity)
            extracted.specification_date = st.text_input("Specification Date", value=extracted.specification_date)
            extracted.storage_temperature = st.text_area(
                "Storage Temperature (from MSDS)", value=extracted.storage_temperature, height=100
            )
            extracted.currency = st.text_input("Default Currency", value=extracted.currency)

        company_info = st.text_area(
            "Company info block (goes at bottom of label and in price list)",
            value=st.session_state.get("company_info", ""),
            height=120,
        )
        st.session_state["company_info"] = company_info
        st.session_state["extracted"] = extracted

    with tab_positions:
        st.write("Edit extracted product rows. Add/remove rows as needed.")
        df = _positions_to_df(extracted.positions)
        edited = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_order=[
                "code",
                "name_en",
                "name_ru",
                "quantity",
                "packing_en",
                "packing_ru",
                "unit_price",
                "total_price",
                "currency",
            ],
        )

        if st.button("Copy EN Name/Packing -> RU (empty cells only)"):
            copied_positions = _df_to_positions(edited)
            for p in copied_positions:
                if not p.name_ru:
                    p.name_ru = p.name_en
                if not p.packing_ru:
                    p.packing_ru = p.packing_en
            extracted.positions = copied_positions
            st.session_state["extracted"] = extracted
            st.success("Copied EN values into empty RU fields. Open Positions tab again to review.")

        if st.button("Apply Position Edits"):
            extracted.positions = _df_to_positions(edited)
            st.session_state["extracted"] = extracted
            st.success("Positions updated.")

    with tab_generate:
        st.write("Generate output documents based on current reviewed fields.")

        if st.button("Build Output Files", type="primary"):
            tpl_price = price_template.getvalue() if price_template else None
            tpl_label = label_template.getvalue() if label_template else None

            price_docx = generate_price_list_doc(
                data=extracted,
                company_info=st.session_state.get("company_info", ""),
                template_bytes=tpl_price,
            )
            label_docx = generate_label_doc(
                data=extracted,
                company_info=st.session_state.get("company_info", ""),
                template_bytes=tpl_label,
            )
            bundle = _pack_outputs(price_docx, label_docx)

            st.download_button(
                "Download Proforma Invoice (.docx)",
                data=price_docx,
                file_name="Proforma_Invoice.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            st.download_button(
                "Download Label (.docx)",
                data=label_docx,
                file_name="Label.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            st.download_button(
                "Download Both (.zip)",
                data=bundle,
                file_name="Customs_Docs.zip",
                mime="application/zip",
            )

    with tab_debug:
        logs = st.session_state.get("logs", {})
        st.json(logs)

        for name, payload in logs.items():
            st.subheader(f"{name.upper()} text preview")
            st.code(payload.get("text_preview", ""), language="text")

else:
    st.info("Upload files and click 'Extract Documents' to begin.")
