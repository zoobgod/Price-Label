"""Extraction pipeline – orchestrates PDF reading and LLM-based extraction."""

from __future__ import annotations

from anthropic import Anthropic

from .llm_extractor import extract_msds, extract_pi, extract_spec
from .models import ExtractedData, Position
from .pdf_reader import pdf_to_native_text, pdf_to_page_images

DEFAULT_STORAGE = "+15C to +25C ambient"


def run_extraction_pipeline(
    pi_pdf_bytes: bytes,
    msds_pdf_bytes: bytes | None,
    specification_pdf_bytes: bytes | None,
    api_key: str,
) -> tuple[ExtractedData, dict]:
    """Run the full extraction pipeline using Claude Vision.

    Returns (ExtractedData, logs_dict).
    """
    client = Anthropic(api_key=api_key)
    logs: dict = {}

    # --- PI ---
    pi_images = pdf_to_page_images(pi_pdf_bytes)
    pi_text = pdf_to_native_text(pi_pdf_bytes)
    pi_raw = extract_pi(client, pi_images, pi_text)
    logs["pi"] = pi_raw

    # --- MSDS ---
    msds_raw: dict = {}
    if msds_pdf_bytes:
        msds_images = pdf_to_page_images(msds_pdf_bytes)
        msds_text = pdf_to_native_text(msds_pdf_bytes)
        msds_raw = extract_msds(client, msds_images, msds_text)
        logs["msds"] = msds_raw

    # --- Specification ---
    spec_raw: dict = {}
    if specification_pdf_bytes:
        spec_images = pdf_to_page_images(specification_pdf_bytes)
        spec_text = pdf_to_native_text(specification_pdf_bytes)
        spec_raw = extract_spec(client, spec_images, spec_text)
        logs["spec"] = spec_raw

    # --- Merge into ExtractedData ---
    currency = pi_raw.get("currency", "")
    storage_temp = msds_raw.get("storage_temperature", "") or DEFAULT_STORAGE

    positions: list[Position] = []
    for pos in pi_raw.get("positions", []):
        positions.append(
            Position(
                code=pos.get("code", ""),
                name_en=pos.get("name_en", ""),
                quantity=pos.get("quantity"),
                packing_en=pos.get("packing", ""),
                unit_price=pos.get("unit_price"),
                total_price=pos.get("total_price"),
                currency=currency,
                storage_temperature=storage_temp,
            )
        )

    if not positions:
        positions = [Position(currency=currency, storage_temperature=storage_temp)]

    data = ExtractedData(
        invoice_no=pi_raw.get("invoice_no", ""),
        invoice_date=pi_raw.get("invoice_date", ""),
        buyer_name=pi_raw.get("buyer_name", ""),
        buyer_address=pi_raw.get("buyer_address", ""),
        exporter_name=pi_raw.get("exporter_name", ""),
        exporter_address=pi_raw.get("exporter_address", ""),
        terms_of_delivery=spec_raw.get("terms_of_delivery") or pi_raw.get("terms_of_delivery", ""),
        period_of_validity=spec_raw.get("period_of_validity", ""),
        specification_date=spec_raw.get("specification_date", ""),
        storage_temperature=storage_temp,
        positions=positions,
        currency=currency,
    )

    return data, logs
