from __future__ import annotations

from dataclasses import asdict
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import streamlit as st

from src.doc_builder import generate_label_docs_by_temperature, generate_price_list_doc
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
                storage_temperature=str(row.get("storage_temperature", "")).strip(),
            )
        )
    return rows


def _pack_outputs(price_docx: bytes, label_docs: dict[str, bytes]) -> bytes:
    bio = BytesIO()
    with ZipFile(bio, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("Price_List.docx", price_docx)
        for filename, content in label_docs.items():
            zf.writestr(filename, content)
    return bio.getvalue()


def _call_pipeline(
    pi_pdf_bytes: bytes,
    msds_pdf_bytes: bytes | None,
    specification_pdf_bytes: bytes | None,
    force_ocr_pi: bool,
    force_ocr_specification: bool,
    force_ocr_msds: bool,
):
    try:
        return run_extraction_pipeline(
            pi_pdf_bytes=pi_pdf_bytes,
            msds_pdf_bytes=msds_pdf_bytes,
            specification_pdf_bytes=specification_pdf_bytes,
            force_ocr_pi=force_ocr_pi,
            force_ocr_specification=force_ocr_specification,
            force_ocr_msds=force_ocr_msds,
        )
    except TypeError:
        return run_extraction_pipeline(
            pi_pdf_bytes=pi_pdf_bytes,
            msds_pdf_bytes=msds_pdf_bytes,
            specification_pdf_bytes=specification_pdf_bytes,
            force_ocr_msds=force_ocr_msds,
        )


def _reset_workflow() -> None:
    for key in ["extracted", "logs", "company_profile", "outputs", "workflow_step"]:
        if key in st.session_state:
            del st.session_state[key]


st.title("Pharmacopeia Customs Document Builder")
st.caption("Step workflow: Extract -> Review -> Generate. Downloads stay available until you reset.")

with st.sidebar:
    st.header("Step 1: Inputs")
    pi_pdf = st.file_uploader("PI / Invoice PDF", type=["pdf"], key="pi_pdf")
    msds_pdf = st.file_uploader("MSDS PDF", type=["pdf"], key="msds_pdf")
    spec_pdf = st.file_uploader("Specification PDF", type=["pdf"], key="spec_pdf")

    st.header("Templates")
    price_template = st.file_uploader("Price List .docx template", type=["docx"], key="price_tpl")
    label_template = st.file_uploader("Label .docx template", type=["docx"], key="label_tpl")

    st.header("OCR")
    force_pi_ocr = st.checkbox("Force OCR on PI/Invoice", value=True)
    force_spec_ocr = st.checkbox("Force OCR on Specification", value=False)
    force_msds_ocr = st.checkbox("Force OCR on MSDS", value=True)

    extract_clicked = st.button("Extract Documents", type="primary", use_container_width=True)
    if st.button("Start From Beginning", use_container_width=True):
        _reset_workflow()
        st.rerun()

if extract_clicked:
    if not pi_pdf:
        st.error("PI/Invoice PDF is required.")
    else:
        with st.spinner("Extracting fields from PDFs..."):
            extracted, logs = _call_pipeline(
                pi_pdf_bytes=pi_pdf.getvalue(),
                msds_pdf_bytes=msds_pdf.getvalue() if msds_pdf else None,
                specification_pdf_bytes=spec_pdf.getvalue() if spec_pdf else None,
                force_ocr_pi=force_pi_ocr,
                force_ocr_specification=force_spec_ocr,
                force_ocr_msds=force_msds_ocr,
            )

        st.session_state["extracted"] = extracted
        st.session_state["logs"] = logs
        st.session_state["workflow_step"] = 2
        st.session_state["outputs"] = {}
        st.session_state["company_profile"] = {
            "exporter_company_name_en": extracted.exporter_name,
            "exporter_company_name_ru": extracted.exporter_name_ru or extracted.exporter_name,
            "exporter_company_address_en": extracted.exporter_address,
            "storage_temperature_ru": "",
        }

        pi_meta = (logs.get("pi") or {}).get("meta") or {}
        if not pi_meta.get("tesseract_available", False):
            st.warning(
                "OCR engine missing. For Streamlit Cloud ensure `packages.txt` contains `tesseract-ocr` and `tesseract-ocr-rus`."
            )

if "extracted" not in st.session_state:
    st.info("Upload files and click 'Extract Documents' to begin.")
else:
    extracted: ExtractedData = st.session_state["extracted"]
    step = st.session_state.get("workflow_step", 2)

    st.subheader(f"Current Step: {step}/3")
    st.progress(step / 3)

    # Step 2: Review
    st.markdown("### Step 2: Review and Correct")
    c1, c2 = st.columns(2)
    with c1:
        extracted.invoice_no = st.text_input("Invoice No", value=extracted.invoice_no)
        extracted.invoice_date = st.text_input("Invoice Date", value=extracted.invoice_date)
        extracted.buyer_name = st.text_input("Buyer Name", value=extracted.buyer_name)
        extracted.buyer_address = st.text_area("Buyer Address", value=extracted.buyer_address, height=100)
        extracted.exporter_name = st.text_input("Exporter Name (EN)", value=extracted.exporter_name)
        extracted.exporter_name_ru = st.text_input(
            "Exporter Name (RU)",
            value=extracted.exporter_name_ru or extracted.exporter_name,
        )

    with c2:
        extracted.terms_of_delivery = st.text_input("Terms Of Delivery", value=extracted.terms_of_delivery)
        extracted.period_of_validity = st.text_input("Period Of Validity", value=extracted.period_of_validity)
        extracted.specification_date = st.text_input("Specification Date", value=extracted.specification_date)
        extracted.storage_temperature = st.text_input(
            "Default Storage Temperature (fallback)",
            value=extracted.storage_temperature,
        )
        extracted.currency = st.text_input("Default Currency", value=extracted.currency)

    profile = st.session_state.get("company_profile", {})
    st.markdown("### Template Company Fields")
    p1, p2 = st.columns(2)
    with p1:
        profile["exporter_company_name_en"] = st.text_input(
            "{{EXPORTER_COMPANY_NAME_EN}}",
            value=profile.get("exporter_company_name_en", extracted.exporter_name),
        )
        profile["exporter_company_name_ru"] = st.text_input(
            "{{EXPORTER_COMPANY_NAME_RU}}",
            value=profile.get("exporter_company_name_ru", extracted.exporter_name_ru or extracted.exporter_name),
        )
    with p2:
        profile["exporter_company_address_en"] = st.text_area(
            "{{EXPORTER_COMPANY_ADRESS_EN}}",
            value=profile.get("exporter_company_address_en", extracted.exporter_address),
            height=90,
        )
        profile["storage_temperature_ru"] = st.text_input(
            "{{STORAGE_TEMPERATURE_RU}} override (optional)",
            value=profile.get("storage_temperature_ru", ""),
        )

    st.session_state["company_profile"] = profile

    st.markdown("### Positions")
    st.caption("Edit rows. `storage_temperature` controls label grouping. Missing temperature defaults to ambient.")
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
            "storage_temperature",
        ],
        key="positions_editor",
    )

    a1, a2, a3 = st.columns(3)
    with a1:
        if st.button("Apply Position Edits"):
            extracted.positions = _df_to_positions(edited)
            st.session_state["extracted"] = extracted
            st.success("Positions updated.")
    with a2:
        if st.button("Copy EN Name/Packing -> RU"):
            copied_positions = _df_to_positions(edited)
            for position in copied_positions:
                if not position.name_ru:
                    position.name_ru = position.name_en
                if not position.packing_ru:
                    position.packing_ru = position.packing_en
                if not position.storage_temperature:
                    position.storage_temperature = extracted.storage_temperature
            extracted.positions = copied_positions
            st.session_state["extracted"] = extracted
            st.success("Copied EN fields and filled empty temperatures from default.")
    with a3:
        if st.button("Proceed to Step 3"):
            st.session_state["workflow_step"] = 3
            st.rerun()

    st.session_state["extracted"] = extracted

    # Step 3: Generate
    if step >= 3:
        st.markdown("### Step 3: Generate Outputs")
        st.caption("Downloads stay visible until you click 'Start From Beginning'.")

        if st.button("Back to Step 2 (Review)"):
            st.session_state["workflow_step"] = 2
            st.rerun()

        if st.button("Generate / Refresh Files", type="primary"):
            tpl_price = price_template.getvalue() if price_template else None
            tpl_label = label_template.getvalue() if label_template else None
            company_profile = st.session_state.get("company_profile", {})

            price_docx = generate_price_list_doc(
                data=extracted,
                company_info="",
                template_bytes=tpl_price,
                company_profile=company_profile,
            )
            label_docs = generate_label_docs_by_temperature(
                data=extracted,
                company_info="",
                template_bytes=tpl_label,
                company_profile=company_profile,
            )
            bundle = _pack_outputs(price_docx, label_docs)

            st.session_state["outputs"] = {
                "price_docx": price_docx,
                "label_docs": label_docs,
                "bundle": bundle,
            }

        outputs = st.session_state.get("outputs") or {}
        if outputs:
            st.success("Generated files are ready.")
            st.download_button(
                "Download Price List (.docx)",
                data=outputs["price_docx"],
                file_name="Price_List.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            for filename, label_bytes in outputs["label_docs"].items():
                st.download_button(
                    f"Download {filename}",
                    data=label_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            st.download_button(
                "Download All Outputs (.zip)",
                data=outputs["bundle"],
                file_name="Customs_Docs.zip",
                mime="application/zip",
            )

    with st.expander("Extraction Debug", expanded=False):
        logs = st.session_state.get("logs", {})
        st.json(logs)
