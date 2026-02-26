from __future__ import annotations

import re
from typing import Iterable

from .models import ExtractedData, Position

DEFAULT_STORAGE_TEMPERATURE_EN = "+15C to +25C ambient"


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


def _extract_numbers(text: str) -> list[float]:
    values: list[float] = []
    for match in re.finditer(r"\b(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\b", text):
        num = _to_float(match.group(0))
        if num is not None:
            values.append(num)
    return values


def _extract_invoice_date(text: str) -> str:
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


def _split_name_and_packing(line: str) -> tuple[str, str]:
    text = _clean_space(line)
    m = re.search(
        r"^(.*?)(\d+\s*(?:mg|g|kg|ml|mcg)(?:\s*/\s*(?:vial|vials|ampoule|ampoules|unit|units|un|шт))?)$",
        text,
        re.IGNORECASE,
    )
    if m:
        return _clean_space(m.group(1)).rstrip("-/"), _clean_space(m.group(2))
    return text, ""


def _find_goods_block_lines(pi_text: str) -> list[str]:
    lines = [ln.strip() for ln in pi_text.splitlines()]
    start = -1
    for i, line in enumerate(lines):
        low = line.lower()
        if "description of goods" in low:
            start = i
            break
    if start == -1:
        return []

    block: list[str] = []
    for line in lines[start + 1 :]:
        low = line.lower()
        if "amount in words" in low or "declaration" in low or low.startswith("for "):
            break
        if line:
            block.append(line)
    return block


def _extract_goods_block_position(pi_text: str, currency: str) -> Position | None:
    block = _find_goods_block_lines(pi_text)
    if not block:
        return None

    block_text = "\n".join(block)

    # Parse numeric triplet (qty, rate, total) close to the end of goods block.
    nums = _extract_numbers(block_text)

    qty: float | None = None
    unit: float | None = None
    total: float | None = None

    if len(nums) >= 3:
        qty, unit, total = nums[-3], nums[-2], nums[-1]
    elif len(nums) == 2:
        qty = 1.0
        unit, total = nums[-2], nums[-1]
    elif len(nums) == 1:
        qty = nums[0]

    header_noise = re.compile(
        r"(mark\s*&\s*nos|part\s*no|description\s*of\s*goods|quantity|rate/item|total\s*value|in\s*vails|in\s*vials|incoterms)",
        re.IGNORECASE,
    )

    name_candidates: list[str] = []
    for line in block:
        if header_noise.search(line):
            continue
        if len(re.findall(r"[A-Za-z]", line)) < 8:
            continue
        if re.fullmatch(r"[\d.,\s]+", line):
            continue
        name_candidates.append(line)

    if not name_candidates:
        return None

    # Keep the most chemical-like / descriptive line.
    name_line = max(name_candidates, key=lambda x: (len(re.findall(r"[A-Za-z]", x)), len(x)))
    name_en, packing = _split_name_and_packing(name_line)

    code = _find_first(
        [
            r"\b([A-Z]{2,}-[A-Z]{2,}-\d{2,4})\b",
            r"\bPART\s*NO\s*[:\-]?\s*([A-Za-z0-9\-/]+)\b",
        ],
        block_text,
    )

    if not code:
        code = "ITEM-1"

    return Position(
        code=code,
        name_en=name_en,
        packing_en=packing,
        quantity=qty,
        unit_price=unit,
        total_price=total,
        currency=currency,
    )


def _extract_positions(pi_text: str, description_text: str, currency: str) -> list[Position]:
    positions: list[Position] = []
    lines = [ln.strip() for ln in pi_text.splitlines() if ln.strip()]

    row_pattern = re.compile(
        r"^(?P<code>[A-Za-z0-9\-/]{3,})\s+(?P<qty>\d+(?:[.,]\d+)?)\s+(?P<unit>[\d,]+(?:\.\d{2})?)\s+(?P<total>[\d,]+(?:\.\d{2})?)$"
    )

    invalid_codes = {"incoterms", "invoice", "payment", "terms", "in"}

    for line in lines:
        normalized = re.sub(r"\s+", " ", line)
        m = row_pattern.match(normalized)
        if not m:
            continue
        code = m.group("code")
        if code.lower() in invalid_codes:
            continue
        positions.append(
            Position(
                code=code,
                quantity=_to_float(m.group("qty")),
                unit_price=_to_float(m.group("unit")),
                total_price=_to_float(m.group("total")),
                currency=currency,
            )
        )

    # Extract names / packing from description-like lines.
    desc_lines = [ln.strip(" -") for ln in description_text.splitlines() if ln.strip()]
    if not desc_lines:
        desc_lines = _find_goods_block_lines(pi_text)

    packing_pattern = re.compile(
        r"\b\d+\s*[xX]\s*\d+\s*[A-Za-z]+\b|\b\d+\s*(?:mg|g|kg|ml|l|mcg|iu)(?:\s*/\s*(?:vial|vials|ampoule|ampoules))?\b",
        re.IGNORECASE,
    )

    name_candidates: list[str] = []
    packing_candidates: list[str] = []

    for line in desc_lines:
        if re.search(r"authorised signatory|amount in words|mark\s*&\s*nos|rate/item|total\s*value|in\s*vails", line, re.IGNORECASE):
            continue
        if packing_pattern.search(line):
            # line can contain both name and packing
            name, pack = _split_name_and_packing(line)
            if pack:
                if re.search(r"^\d+\s*[xX]\s*\d+", line):
                    packing_candidates.append(_clean_space(line))
                else:
                    packing_candidates.append(pack)
                if len(re.findall(r"[A-Za-z]", name)) >= 4:
                    name_candidates.append(name)
            else:
                packing_candidates.append(_clean_space(line))
        elif len(re.findall(r"[A-Za-z]", line)) >= 8:
            name_candidates.append(_clean_space(line))

    if not positions:
        fallback = _extract_goods_block_position(pi_text, currency)
        if fallback:
            positions = [fallback]

    if not positions:
        positions = [Position(code="ITEM-1", currency=currency)]

    for idx, pos in enumerate(positions):
        if idx < len(name_candidates) and not pos.name_en:
            pos.name_en = name_candidates[idx]
        elif name_candidates and not pos.name_en:
            pos.name_en = name_candidates[0]

        if idx < len(packing_candidates) and not pos.packing_en:
            pos.packing_en = packing_candidates[idx]
        elif packing_candidates and not pos.packing_en:
            pos.packing_en = packing_candidates[0]

        # Last-resort cleanup for lines that still combine name+packing.
        if pos.name_en and not pos.packing_en:
            name, pack = _split_name_and_packing(pos.name_en)
            if pack:
                pos.name_en = name
                pos.packing_en = pack

    # If parsed rows look suspicious (no prices), prefer goods-block structured row.
    if positions and (positions[0].unit_price is None or positions[0].total_price is None):
        fallback = _extract_goods_block_position(pi_text, currency)
        if fallback and fallback.unit_price is not None and fallback.total_price is not None:
            if not fallback.name_en and positions[0].name_en:
                fallback.name_en = positions[0].name_en
            if not fallback.packing_en and positions[0].packing_en:
                fallback.packing_en = positions[0].packing_en
            positions[0] = fallback

    return positions


def _extract_description_lines(pi_text: str) -> list[str]:
    lines = [ln.strip() for ln in pi_text.splitlines()]
    start_idx = -1
    for idx, line in enumerate(lines):
        if "description of goods" in line.lower():
            start_idx = idx
            break
    if start_idx == -1:
        return []

    collected: list[str] = []
    for line in lines[start_idx + 1 :]:
        low = line.lower()
        if "amount in words" in low or "declaration" in low:
            break
        if line.strip():
            collected.append(line.strip())
    return collected


def _normalize_terms(value: str) -> str:
    if not value:
        return ""
    text = _clean_space(value)
    low = text.lower()
    incoterm = _find_first([r"\b(CPT|FOB|CIF|EXW|DAP|DDP|FCA)\b"], text)
    if not incoterm:
        return text

    has_air = bool(re.search(r"\b(by\s+air|air)\b", low))
    city = ""
    if "moscow" in low:
        city = "MOSCOW"
    elif "hyderabad" in low:
        city = "HYDERABAD"

    result = incoterm.upper()
    if has_air:
        result += " BY AIR"
    if city:
        result += f" {city}"
    return result


def _extract_terms(pi_text: str) -> str:
    lines = [ln.strip() for ln in pi_text.splitlines() if ln.strip()]

    # First pass: direct incoterm lines.
    for line in lines:
        if re.search(r"\b(CPT|FOB|CIF|EXW|DAP|DDP|FCA)\b", line, re.IGNORECASE):
            normalized = _normalize_terms(line)
            if normalized:
                return normalized

    # Second pass: 'incoterms' label lines.
    for line in lines:
        if "incoterms" in line.lower():
            normalized = _normalize_terms(line)
            if normalized:
                return normalized

    # Third pass: terms header with next-line lookup.
    for idx, line in enumerate(lines):
        if "terms of delivery" in line.lower() or "delivery conditions" in line.lower():
            window = lines[idx : idx + 6]
            for candidate in window:
                normalized = _normalize_terms(candidate)
                if normalized and normalized != _clean_space(candidate):
                    return normalized
    return ""


def parse_pi_text(pi_text: str) -> ExtractedData:
    currency = _extract_currency(pi_text)

    invoice_no = _find_first(
        [
            r"\b([A-Z]{2,}/PI/\d{2}-\d{2}/\d+)\b",
            r"\b([A-Z]{2,}[/-][A-Z][/-]\d{2}-\d{2}[/-]\d+)\b",
            r"Invoice\s*No\.?\s*&?\s*Date\s*\n\s*([A-Za-z0-9\-/]+)",
        ],
        pi_text,
    )

    exporter_name = _find_first(
        [
            r"Exporter:\s*\n\s*([^\n]+)",
            r"\b(M/S\.[^\n]+?)\s+[A-Z]{2,}/PI/\d{2}-\d{2}/\d+\b",
        ],
        pi_text,
    )
    exporter_address = _extract_block(
        pi_text,
        start_label="Exporter:",
        stop_labels=["Invoice No", "Consignee", "Buyer", "IEC NO", "GSTIN"],
    )

    buyer_block = _extract_block(
        pi_text,
        start_label="Consignee:",
        stop_labels=["Buyer", "Country of Origin", "Pre-Carriage", "Terms of Delivery"],
    )
    buyer_lines = [ln for ln in buyer_block.splitlines() if ln.strip()]
    buyer_name = buyer_lines[0] if buyer_lines else ""
    buyer_address = "\n".join(buyer_lines[1:]).strip() if len(buyer_lines) > 1 else ""

    description_lines = _extract_description_lines(pi_text)
    description = "\n".join(description_lines)

    positions = _extract_positions(pi_text=pi_text, description_text=description, currency=currency)

    data = ExtractedData(
        invoice_no=invoice_no,
        invoice_date=_extract_invoice_date(pi_text),
        buyer_name=buyer_name,
        buyer_address=buyer_address,
        exporter_name=exporter_name,
        exporter_address=exporter_address,
        terms_of_delivery=_extract_terms(pi_text),
        specification_date=_find_first(
            [
                r"Specification\s*No[^\n]*DT[:\s]*([0-9./-]+)",
                r"Contract\s*No[^\n]*Dt[:\s]*([0-9./-]+)",
            ],
            pi_text,
        ),
        positions=positions,
        currency=currency,
    )

    return data


def parse_specification_text(spec_text: str) -> dict[str, str]:
    result: dict[str, str] = {}

    terms_raw = _find_first(
        [
            r"Delivery\s*conditions\s*[:\-]?\s*([^\n.]+)",
            r"Terms\s*of\s*Delivery\s*[:\-]?\s*([^\n.]+)",
            r"Условия\s*поставки\s*[:\-]?\s*([^\n.]+)",
        ],
        spec_text,
    )
    result["terms_of_delivery"] = _normalize_terms(terms_raw)

    result["period_of_validity"] = _find_first(
        [
            r"Shipment\s*time\s*[:\-]?\s*([^\n.]+)",
            r"Period\s*of\s*Validity\s*[:\-]?\s*([^\n.]+)",
            r"Validity\s*Period\s*[:\-]?\s*([^\n.]+)",
            r"Valid\s*for\s*[:\-]?\s*([^\n.]+)",
        ],
        spec_text,
    )

    result["specification_date"] = _find_first(
        [
            r"Specification\s*\S*\s*of\s*(\d{2}[./-]\d{2}[./-]\d{4})",
            r"Moscow\s*/\s*Москва\s*(\d{2}[./-]\d{2}[./-]\d{4})",
            r"(\d{2}[./-]\d{2}[./-]\d{4})",
        ],
        spec_text,
    )

    result["packing"] = _find_first(
        [
            r"Packing\s*[:\-]?\s*([^\n.]+)",
            r"Упаковка\s*[:\-]?\s*([^\n.]+)",
        ],
        spec_text,
    )

    return result


def _normalize_storage_temperature(text: str) -> str:
    if not text:
        return ""

    low = text.lower()
    if "room temperature" in low or "ambient" in low:
        return DEFAULT_STORAGE_TEMPERATURE_EN

    m = re.search(
        r"(?:between|from|at)?\s*([+-]?\d{1,2})\s*(?:°?\s*[cfCF])\s*(?:to|and|–|-|~)\s*([+-]?\d{1,2})\s*(?:°?\s*[cfCF])",
        text,
        re.IGNORECASE,
    )
    if m:
        left = int(m.group(1))
        right = int(m.group(2))

        def _fmt(v: int) -> str:
            return f"+{v}" if v >= 0 else str(v)

        return f"{_fmt(left)}C to {_fmt(right)}C"

    return ""


def parse_msds_text(msds_text: str) -> dict[str, str]:
    lines = [ln.strip() for ln in msds_text.splitlines() if ln.strip()]

    for line in lines:
        if re.search(r"storage|store|shipping|transport|keep|maintain", line, re.IGNORECASE):
            normalized = _normalize_storage_temperature(line)
            if normalized:
                return {"storage_temperature": normalized}

    temp_match = re.search(
        r"(?:between|from|at)\s*([+-]?\d{1,2})\s*(?:°?\s*[CFcf])\s*(?:to|and|–|-|~)\s*([+-]?\d{1,2})\s*(?:°?\s*[CFcf])",
        msds_text,
        re.IGNORECASE,
    )
    if temp_match:
        return {"storage_temperature": _normalize_storage_temperature(temp_match.group(0))}

    if re.search(r"room temperature|ambient", msds_text, re.IGNORECASE):
        return {"storage_temperature": DEFAULT_STORAGE_TEMPERATURE_EN}

    return {"storage_temperature": ""}


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
    if not pi_data.storage_temperature:
        pi_data.storage_temperature = DEFAULT_STORAGE_TEMPERATURE_EN

    packing_from_spec = spec_data.get("packing", "")

    for position in pi_data.positions:
        if not position.currency:
            position.currency = pi_data.currency
        if not position.storage_temperature:
            position.storage_temperature = pi_data.storage_temperature
        if not position.packing_en and packing_from_spec:
            position.packing_en = packing_from_spec

    return pi_data
