from __future__ import annotations

import re
from typing import Iterable

from .models import ExtractedData, Position


def _clean_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _find_first(patterns: Iterable[str], text: str, flags: int = re.IGNORECASE) -> str:
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return _clean_space(m.group(1))
    return ""


def _extract_currency(text: str) -> str:
    if "₹" in text or re.search(r"\bINR\b", text, re.IGNORECASE):
        return "INR"
    if "$" in text or re.search(r"\bUSD\b", text, re.IGNORECASE):
        return "USD"
    if "€" in text or re.search(r"\bEUR\b", text, re.IGNORECASE):
        return "EUR"

    m = re.search(r"\(In\s+([A-Z]{3})\)", text)
    if m:
        return m.group(1)
    return ""


def _to_float(value: str) -> float | None:
    candidate = value.replace(" ", "").replace(",", "").strip()
    candidate = re.sub(r"[^\d.\-]", "", candidate)
    if not candidate:
        return None
    try:
        return float(candidate)
    except ValueError:
        return None


def _extract_invoice_date(text: str) -> str:
    # Common forms: 26-Feb-26, 26.02.2026, 26/02/2026
    m = re.search(r"\b(\d{1,2}[-/.][A-Za-z]{3,9}[-/.]\d{2,4})\b", text)
    if m:
        return m.group(1)
    m = re.search(r"\b\d{1,2}([./-])\d{1,2}\1\d{2,4}\b", text)
    return m.group(0) if m else ""


def _extract_block(text: str, start_label: str, stop_labels: list[str]) -> str:
    lines = [ln.strip() for ln in text.splitlines()]
    start_idx = -1
    for idx, line in enumerate(lines):
        if start_label.lower() in line.lower():
            start_idx = idx
            break
    if start_idx == -1:
        return ""

    collected: list[str] = []
    for idx in range(start_idx + 1, len(lines)):
        line = lines[idx]
        if not line:
            continue
        if any(stop.lower() in line.lower() for stop in stop_labels):
            break
        collected.append(line)
    return "\n".join(collected).strip()


def _extract_positions(pi_text: str, description_text: str, currency: str) -> list[Position]:
    positions: list[Position] = []
    lines = [ln.strip() for ln in pi_text.splitlines() if ln.strip()]

    row_pattern = re.compile(
        r"^(?P<code>[A-Za-z0-9\-/]{3,})\s+(?P<qty>\d+(?:[.,]\d+)?)\s+(?P<unit>[\d,]+(?:\.\d{2})?)\s+(?P<total>[\d,]+(?:\.\d{2})?)$"
    )
    row_pattern_loose = re.compile(
        r"^(?P<code>[A-Za-z0-9\-/]{3,})\s+(?P<qty>\d+(?:[.,]\d+)?)\s+(?P<unit>[0-9,.\s]{4,20})\s+(?P<total>[0-9,.\s]{4,20})$"
    )

    for line in lines:
        normalized = re.sub(r"\s+", " ", line)
        m = row_pattern.match(normalized) or row_pattern_loose.match(normalized)
        if not m:
            continue
        positions.append(
            Position(
                code=m.group("code"),
                quantity=_to_float(m.group("qty")),
                unit_price=_to_float(m.group("unit")),
                total_price=_to_float(m.group("total")),
                currency=currency,
            )
        )

    if not positions:
        # Fallback for OCR output where prices are split or missing.
        code_qty_pattern = re.compile(r"^(?P<code>[A-Za-z0-9\-/]{3,})\s+(?P<qty>\d+(?:[.,]\d+)?)\b")
        for line in lines:
            m = code_qty_pattern.match(re.sub(r"\s+", " ", line))
            if not m:
                continue
            if any(p.code == m.group("code") for p in positions):
                continue
            positions.append(
                Position(
                    code=m.group("code"),
                    quantity=_to_float(m.group("qty")),
                    currency=currency,
                )
            )

    # Extract potential packings and names from description block.
    desc_lines = [ln.strip(" -") for ln in description_text.splitlines() if ln.strip()]
    packing_candidates: list[str] = []
    name_candidates: list[str] = []
    packing_pattern = re.compile(
        r"\b\d+\s*[xX]\s*\d+\s*[A-Za-z]+\b|\b\d+\s*(?:mg|g|kg|ml|l|mcg|iu)\b",
        re.IGNORECASE,
    )

    for ln in desc_lines:
        if packing_pattern.search(ln):
            packing_candidates.append(ln)
        elif len(ln.split()) >= 2 and not re.search(r"amount in words|authorised signatory", ln, re.IGNORECASE):
            name_candidates.append(ln)

    if not name_candidates:
        # OCR fallback: select long alpha-heavy lines as probable names.
        for ln in lines:
            if re.search(r"(invoice|exporter|consignee|quantity|total|declaration|authorised)", ln, re.IGNORECASE):
                continue
            if len(re.findall(r"[A-Za-z]", ln)) >= 12 and len(ln.split()) <= 8:
                name_candidates.append(ln)
        name_candidates = name_candidates[:3]

    if not positions:
        positions = [Position(currency=currency)]

    for idx, pos in enumerate(positions):
        if idx < len(name_candidates):
            pos.name_en = name_candidates[idx]
        elif name_candidates:
            pos.name_en = name_candidates[0]

        if idx < len(packing_candidates):
            pos.packing_en = packing_candidates[idx]
        elif packing_candidates:
            pos.packing_en = packing_candidates[0]

    return positions


def _extract_description_lines(pi_text: str) -> list[str]:
    lines = [ln.strip() for ln in pi_text.splitlines()]
    start_idx = -1
    for idx, ln in enumerate(lines):
        if "description of goods" in ln.lower():
            start_idx = idx
            break
    if start_idx == -1:
        return []

    collected: list[str] = []
    for idx in range(start_idx + 1, len(lines)):
        ln = lines[idx].strip()
        if not ln:
            continue
        low = ln.lower()
        if "amount in words" in low or low.startswith("for "):
            break
        if "authorised signatory" in low:
            continue
        collected.append(ln)
    return collected


def _extract_terms(pi_text: str) -> str:
    lines = [ln.strip() for ln in pi_text.splitlines() if ln.strip()]
    incoterm_re = re.compile(r"\b(CPT|FOB|CIF|EXW|DAP|DDP|FCA)\b", re.IGNORECASE)
    for ln in lines:
        if "incoterms" in ln.lower():
            return _clean_space(ln)
    for ln in lines:
        if incoterm_re.search(ln):
            return _clean_space(ln)
    # OCR fallback: sometimes "Terms of Delivery and Payment" is recognized
    # but incoterm is on the next line.
    for idx, ln in enumerate(lines[:-1]):
        if "terms" in ln.lower() and "delivery" in ln.lower():
            nxt = lines[idx + 1]
            if len(nxt.split()) > 2:
                return _clean_space(nxt)
    return ""


def parse_pi_text(pi_text: str) -> ExtractedData:
    currency = _extract_currency(pi_text)
    invoice_no = _find_first(
        [
            r"\b([A-Z]{2,}[/-][A-Z][/-]\d{2}-\d{2}[/-]\d+)\b",
            r"Invoice\s*No\.?\s*&?\s*Date\s*\n\s*([A-Za-z0-9\-/]+)",
            r"Invoice\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\-/]+)",
        ],
        pi_text,
    )

    exporter_name = _find_first([r"\b(M/S\.[^\n]+?)\s+[A-Z]{2,}[/-][A-Z][/-]\d{2}-\d{2}[/-]\d+\b"], pi_text)
    if not exporter_name:
        exporter_name = _find_first([r"Exporter:\s*([^\n]+)"], pi_text)
    exporter_address = _extract_block(
        pi_text,
        start_label="Exporter:",
        stop_labels=["Consignee", "Buyer", "GSTIN", "IEC NO"],
    )

    buyer_block = _extract_block(
        pi_text,
        start_label="Consignee:",
        stop_labels=["Buyer", "Pre-Carriage", "Vessel/Flight", "Quantity", "Port of"],
    )
    buyer_lines = buyer_block.splitlines()
    buyer_name = buyer_lines[0] if buyer_lines else ""
    buyer_address = "\n".join(buyer_lines[1:]).strip() if len(buyer_lines) > 1 else ""

    description_lines = _extract_description_lines(pi_text)
    description = "\n".join(description_lines)

    terms = _extract_terms(pi_text)

    positions = _extract_positions(pi_text=pi_text, description_text=description, currency=currency)

    data = ExtractedData(
        invoice_no=invoice_no,
        invoice_date=_extract_invoice_date(pi_text),
        buyer_name=buyer_name,
        buyer_address=buyer_address,
        exporter_name=exporter_name,
        exporter_address=exporter_address,
        terms_of_delivery=terms,
        specification_date=_find_first([r"Specification\s*No[^\n]*DT:\s*([0-9./-]+)"], pi_text),
        positions=positions,
        currency=currency,
    )

    if len(positions) == 1 and not positions[0].name_en:
        name_candidate = _find_first([r"Description\s*of\s*Goods\s*\n([^\n]+)"], pi_text)
        if name_candidate:
            data.positions[0].name_en = name_candidate

    return data


def parse_specification_text(spec_text: str) -> dict[str, str]:
    result: dict[str, str] = {}

    result["terms_of_delivery"] = _find_first(
        [
            r"Terms\s*of\s*Delivery\s*[:\-]?\s*([^\n]+)",
            r"Delivery\s*Terms\s*[:\-]?\s*([^\n]+)",
        ],
        spec_text,
    )

    result["period_of_validity"] = _find_first(
        [
            r"Period\s*of\s*Validity\s*[:\-]?\s*([^\n]+)",
            r"Validity\s*Period\s*[:\-]?\s*([^\n]+)",
            r"Valid\s*for\s*[:\-]?\s*([^\n]+)",
        ],
        spec_text,
    )

    result["specification_date"] = _find_first(
        [
            r"Specification\s*Date\s*[:\-]?\s*([^\n]+)",
            r"Date\s*of\s*Specification\s*[:\-]?\s*([^\n]+)",
            r"Spec\.?\s*Date\s*[:\-]?\s*([^\n]+)",
        ],
        spec_text,
    )

    return result


def parse_msds_text(msds_text: str) -> dict[str, str]:
    lines = [ln.strip() for ln in msds_text.splitlines() if ln.strip()]

    # Keep full line if it contains both storage keyword and temperature indicator.
    for ln in lines:
        if re.search(r"storage|store", ln, re.IGNORECASE) and re.search(r"°|deg|c\b|f\b|below|between|ambient|room", ln, re.IGNORECASE):
            return {"storage_temperature": _clean_space(ln)}

    # Shipping/handling condition phrasing from OCR text.
    for ln in lines:
        if re.search(r"shipping|keep|maintain", ln, re.IGNORECASE) and re.search(
            r"between\s*\(?\s*-?\d+\s*[°]?\s*[CF]?\s*[–\-to]+\s*-?\d+\s*[°]?\s*[CF]?",
            ln,
            re.IGNORECASE,
        ):
            return {"storage_temperature": _clean_space(ln)}

    # Fallback: explicit range from nearby context.
    temp_match = re.search(
        r"(-?\d+\s*(?:to|–|-)\s*-?\d+\s*(?:°\s*[CF]|degrees?\s*[CF]|[CF]))",
        msds_text,
        re.IGNORECASE,
    )
    if temp_match:
        return {"storage_temperature": f"Store at {temp_match.group(1)}"}

    ctx = _find_first(
        [
            r"Storage\s*(?:conditions?|temperature)?\s*[:\-]?\s*([^\n]+)",
            r"Recommended\s*storage\s*[:\-]?\s*([^\n]+)",
        ],
        msds_text,
    )
    return {"storage_temperature": ctx}


def merge_extraction(
    pi_data: ExtractedData,
    spec_data: dict[str, str] | None,
    msds_data: dict[str, str] | None,
) -> ExtractedData:
    spec_data = spec_data or {}
    msds_data = msds_data or {}

    pi_data.terms_of_delivery = spec_data.get("terms_of_delivery") or pi_data.terms_of_delivery
    pi_data.period_of_validity = spec_data.get("period_of_validity") or pi_data.period_of_validity
    pi_data.specification_date = spec_data.get("specification_date") or pi_data.specification_date
    pi_data.storage_temperature = msds_data.get("storage_temperature") or pi_data.storage_temperature

    for position in pi_data.positions:
        if not position.currency:
            position.currency = pi_data.currency

    return pi_data
