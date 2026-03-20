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

# ── helpers ──────────────────────────────────────────────────────────────────


def _positions_to_df(positions: list[Position]) -> pd.DataFrame:
    records = [asdict(p) for p in positions] if positions else [asdict(Position())]
    return pd.DataFrame(records)


def _df_to_positions(df: pd.DataFrame) -> list[Position]:
    rows: list[Position] = []
    for _, row in df.fillna("").iterrows():

        def _f(v):
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
                quantity=_f(row.get("quantity", "")),
                packing_en=str(row.get("packing_en", "")).strip(),
                packing_ru=str(row.get("packing_ru", "")).strip(),
                unit_price=_f(row.get("unit_price", "")),
                total_price=_f(row.get("total_price", "")),
                currency=str(row.get("currency", "")).strip(),
                storage_temperature=str(row.get("storage_temperature", "")).strip(),
            )
        )
    return rows


def _pack_zip(price_docx: bytes, label_docs: dict[str, bytes]) -> bytes:
    bio = BytesIO()
    with ZipFile(bio, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("Price_List.docx", price_docx)
        for name, content in label_docs.items():
            zf.writestr(name, content)
    return bio.getvalue()


def _reset():
    for key in ("extracted", "logs", "company_profile", "outputs", "step"):
        st.session_state.pop(key, None)


# ── sidebar ──────────────────────────────────────────────────────────────────

st.title("Pharmacopeia Customs Document Builder")
st.caption("Upload PDFs → Extract via Claude Vision → Review → Generate Price List + Labels")

with st.sidebar:
    st.header("API Key")
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        help="Required for extraction. On Streamlit Cloud you can also set ANTHROPIC_API_KEY in Secrets.",
    )
    if not api_key:
        api_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""

    st.header("Input PDFs")
    pi_pdf = st.file_uploader("PI / Invoice PDF", type=["pdf"], key="pi_pdf")
    msds_pdf = st.file_uploader("MSDS PDF (optional)", type=["pdf"], key="msds_pdf")
    spec_pdf = st.file_uploader("Specification PDF (optional)", type=["pdf"], key="spec_pdf")

    st.header("Templates (.docx)")
    price_tpl = st.file_uploader("Price List template", type=["docx"], key="price_tpl")
    label_tpl = st.file_uploader("Label template", type=["docx"], key="label_tpl")

    st.divider()
    extract_btn = st.button("Extract Documents", type="primary", use_container_width=True)
    if st.button("Start From Beginning", use_container_width=True):
        _reset()
        st.rerun()

# ── step 1: extract ─────────────────────────────────────────────────────────

if extract_btn:
    if not pi_pdf:
        st.error("PI / Invoice PDF is required.")
    elif not api_key:
        st.error("Please enter your Anthropic API key in the sidebar.")
    else:
        with st.spinner("Sending PDFs to Claude for extraction... (this may take 30-60 s)"):
            try:
                extracted, logs = run_extraction_pipeline(
                    pi_pdf_bytes=pi_pdf.getvalue(),
                    msds_pdf_bytes=msds_pdf.getvalue() if msds_pdf else None,
                    specification_pdf_bytes=spec_pdf.getvalue() if spec_pdf else None,
                    api_key=api_key,
                )
                st.session_state["extracted"] = extracted
                st.session_state["logs"] = logs
                st.session_state["step"] = 2
                st.session_state.pop("outputs", None)
                st.session_state["company_profile"] = {
                    "exporter_company_name_en": extracted.exporter_name,
                    "exporter_company_name_ru": extracted.exporter_name_ru or extracted.exporter_name,
                    "exporter_company_address_en": extracted.exporter_address,
                    "storage_temperature_ru": "",
                }
            except Exception as exc:
                st.error(f"Extraction failed: {exc}")

# ── step 2: review ──────────────────────────────────────────────────────────

if "extracted" not in st.session_state:
    st.info("Upload files and click **Extract Documents** to begin.")
    st.stop()

extracted: ExtractedData = st.session_state["extracted"]
step = st.session_state.get("step", 2)

st.subheader(f"Step {step} of 3")
st.progress(step / 3)

st.markdown("### Step 2 — Review & Correct Extracted Data")

c1, c2 = st.columns(2)
with c1:
    extracted.invoice_no = st.text_input("Invoice No", value=extracted.invoice_no)
    extracted.invoice_date = st.text_input("Invoice Date (DD.MM.YY)", value=extracted.invoice_date)
    extracted.buyer_name = st.text_input("Buyer Name", value=extracted.buyer_name)
    extracted.buyer_address = st.text_area("Buyer Address", value=extracted.buyer_address, height=80)
    extracted.exporter_name = st.text_input("Exporter Name (EN)", value=extracted.exporter_name)
    extracted.exporter_name_ru = st.text_input(
        "Exporter Name (RU)",
        value=extracted.exporter_name_ru or extracted.exporter_name,
    )

with c2:
    extracted.exporter_address = st.text_area("Exporter Address", value=extracted.exporter_address, height=80)
    extracted.terms_of_delivery = st.text_input("Terms Of Delivery", value=extracted.terms_of_delivery)
    extracted.period_of_validity = st.text_input("Period Of Validity", value=extracted.period_of_validity)
    extracted.specification_date = st.text_input("Specification Date", value=extracted.specification_date)
    extracted.storage_temperature = st.text_input(
        "Default Storage Temperature",
        value=extracted.storage_temperature,
    )
    extracted.currency = st.text_input("Currency", value=extracted.currency)

# Company profile fields (for label template)
profile = st.session_state.get("company_profile", {})
with st.expander("Template Company Fields", expanded=False):
    p1, p2 = st.columns(2)
    with p1:
        profile["exporter_company_name_en"] = st.text_input(
            "EXPORTER_COMPANY_NAME_EN",
            value=profile.get("exporter_company_name_en", extracted.exporter_name),
        )
        profile["exporter_company_name_ru"] = st.text_input(
            "EXPORTER_COMPANY_NAME_RU",
            value=profile.get("exporter_company_name_ru", extracted.exporter_name),
        )
    with p2:
        profile["exporter_company_address_en"] = st.text_area(
            "EXPORTER_COMPANY_ADRESS_EN",
            value=profile.get("exporter_company_address_en", extracted.exporter_address),
            height=80,
        )
        profile["storage_temperature_ru"] = st.text_input(
            "STORAGE_TEMPERATURE_RU override",
            value=profile.get("storage_temperature_ru", ""),
        )
    st.session_state["company_profile"] = profile

# Positions table
st.markdown("### Positions")
st.caption("Edit rows directly. `storage_temperature` per position controls label grouping.")
df = _positions_to_df(extracted.positions)
edited = st.data_editor(
    df,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_order=[
        "code", "name_en", "name_ru", "quantity",
        "packing_en", "packing_ru", "unit_price", "total_price",
        "currency", "storage_temperature",
    ],
    key="positions_editor",
)

b1, b2, b3 = st.columns(3)
with b1:
    if st.button("Apply Edits"):
        extracted.positions = _df_to_positions(edited)
        st.session_state["extracted"] = extracted
        st.success("Positions updated.")
with b2:
    if st.button("Copy EN → RU (name & packing)"):
        positions = _df_to_positions(edited)
        for p in positions:
            if not p.name_ru:
                p.name_ru = p.name_en
            if not p.packing_ru:
                p.packing_ru = p.packing_en
            if not p.storage_temperature:
                p.storage_temperature = extracted.storage_temperature
        extracted.positions = positions
        st.session_state["extracted"] = extracted
        st.success("Copied.")
with b3:
    if st.button("Proceed to Step 3 →", type="primary"):
        extracted.positions = _df_to_positions(edited)
        st.session_state["extracted"] = extracted
        st.session_state["step"] = 3
        st.rerun()

st.session_state["extracted"] = extracted

# ── step 3: generate ────────────────────────────────────────────────────────

if step >= 3:
    st.markdown("---")
    st.markdown("### Step 3 — Generate & Download")

    if st.button("← Back to Review"):
        st.session_state["step"] = 2
        st.rerun()

    if st.button("Generate Files", type="primary"):
        tpl_price = price_tpl.getvalue() if price_tpl else None
        tpl_label = label_tpl.getvalue() if label_tpl else None
        cp = st.session_state.get("company_profile", {})

        price_docx = generate_price_list_doc(
            data=extracted,
            template_bytes=tpl_price,
            company_profile=cp,
        )
        label_docs = generate_label_docs_by_temperature(
            data=extracted,
            template_bytes=tpl_label,
            company_profile=cp,
        )
        bundle = _pack_zip(price_docx, label_docs)

        st.session_state["outputs"] = {
            "price_docx": price_docx,
            "label_docs": label_docs,
            "bundle": bundle,
        }

    outputs = st.session_state.get("outputs")
    if outputs:
        st.success("Files ready for download.")
        st.download_button(
            "Download Price List (.docx)",
            data=outputs["price_docx"],
            file_name="Price_List.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        for fname, lbl_bytes in outputs["label_docs"].items():
            st.download_button(
                f"Download {fname}",
                data=lbl_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        st.download_button(
            "Download All (.zip)",
            data=outputs["bundle"],
            file_name="Customs_Docs.zip",
            mime="application/zip",
        )

# ── debug ────────────────────────────────────────────────────────────────────

with st.expander("Extraction Debug Log", expanded=False):
    st.json(st.session_state.get("logs", {}))
