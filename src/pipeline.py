from __future__ import annotations

from dataclasses import asdict

from .extractors import merge_extraction, parse_msds_text, parse_pi_text, parse_specification_text
from .models import ExtractedData
from .pdf_reader import extract_pdf_text


def _pi_parse_score(data: ExtractedData) -> int:
    score = 0
    if data.invoice_no:
        score += 3
    if data.invoice_date:
        score += 2
    if data.buyer_name:
        score += 1
    if data.terms_of_delivery:
        score += 2
    if data.positions:
        score += 2
        first = data.positions[0]
        if first.code:
            score += 2
            if " " in first.code:
                score -= 1
            if first.code.lower().startswith(("tin", "mfg", "exp")):
                score -= 3
        if first.unit_price is not None:
            score += 2
        if first.total_price is not None:
            score += 2
        if first.name_en and len(first.name_en) > 8:
            score += 1
    return score


def run_extraction_pipeline(
    pi_pdf_bytes: bytes,
    msds_pdf_bytes: bytes | None,
    specification_pdf_bytes: bytes | None,
    force_ocr_pi: bool = True,
    force_ocr_specification: bool = False,
    force_ocr_msds: bool = True,
    **_compat_kwargs: object,
) -> tuple[ExtractedData, dict]:
    logs: dict = {}

    pi_text_native = extract_pdf_text(pi_pdf_bytes, force_ocr=False, prefer_ocr=False)
    pi_data_native = parse_pi_text(pi_text_native.text)
    native_score = _pi_parse_score(pi_data_native)

    pi_data = pi_data_native
    pi_text = pi_text_native
    selected_source = "native"
    ocr_score = None

    if force_ocr_pi:
        pi_text_ocr = extract_pdf_text(pi_pdf_bytes, force_ocr=True, prefer_ocr=True)
        pi_data_ocr = parse_pi_text(pi_text_ocr.text)
        ocr_score = _pi_parse_score(pi_data_ocr)
        if ocr_score > native_score:
            pi_data = pi_data_ocr
            pi_text = pi_text_ocr
            selected_source = "ocr"
        logs["pi_ocr_candidate"] = {"meta": asdict(pi_text_ocr.meta), "text_preview": pi_text_ocr.text[:3000]}

    logs["pi"] = {
        "selected_source": selected_source,
        "native_score": native_score,
        "ocr_score": ocr_score,
        "meta": asdict(pi_text.meta),
        "text_preview": pi_text.text[:3000],
    }

    spec_data = {}
    if specification_pdf_bytes:
        spec_text = extract_pdf_text(specification_pdf_bytes, force_ocr=force_ocr_specification, prefer_ocr=True)
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
