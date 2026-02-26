from __future__ import annotations

from dataclasses import asdict

from .extractors import merge_extraction, parse_msds_text, parse_pi_text, parse_specification_text
from .models import ExtractedData
from .pdf_reader import extract_pdf_text


def run_extraction_pipeline(
    pi_pdf_bytes: bytes,
    msds_pdf_bytes: bytes | None,
    specification_pdf_bytes: bytes | None,
    force_ocr_msds: bool = True,
) -> tuple[ExtractedData, dict]:
    logs: dict = {}

    pi_text = extract_pdf_text(pi_pdf_bytes, force_ocr=False)
    logs["pi"] = {"meta": asdict(pi_text.meta), "text_preview": pi_text.text[:3000]}
    pi_data = parse_pi_text(pi_text.text)

    spec_data = {}
    if specification_pdf_bytes:
        spec_text = extract_pdf_text(specification_pdf_bytes, force_ocr=False)
        logs["specification"] = {"meta": asdict(spec_text.meta), "text_preview": spec_text.text[:3000]}
        spec_data = parse_specification_text(spec_text.text)

    msds_data = {}
    if msds_pdf_bytes:
        msds_text = extract_pdf_text(msds_pdf_bytes, force_ocr=force_ocr_msds)
        logs["msds"] = {"meta": asdict(msds_text.meta), "text_preview": msds_text.text[:3000]}
        msds_data = parse_msds_text(msds_text.text)

    merged = merge_extraction(pi_data=pi_data, spec_data=spec_data, msds_data=msds_data)
    merged.raw = {
        "spec_data": spec_data,
        "msds_data": msds_data,
    }

    return merged, logs
