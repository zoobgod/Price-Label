from __future__ import annotations

import re
from collections import OrderedDict
from datetime import datetime
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from .models import ExtractedData, Position

DEFAULT_STORAGE_TEMPERATURE_EN = "+15C to +25C ambient"
DEFAULT_STORAGE_TEMPERATURE_RU = "+15C до +25C"


def _set_default_font(document: Document, font_name: str = "Times New Roman", size_pt: int = 12) -> None:
    style = document.styles["Normal"]
    style.font.name = font_name
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    style.font.size = Pt(size_pt)


def _clean_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _price_str(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:,.2f}"


def _qty_str(value: float | None) -> str:
    if value is None:
        return ""
    if float(value).is_integer():
        return str(int(value))
    return f"{value:g}"


def _money_with_currency(value: float | None, currency: str) -> str:
    money = _price_str(value)
    if not money:
        return ""
    return f"{money} {currency}".strip()


def _normalize_temp(temp: str) -> str:
    t = _clean_space(temp)
    if not t:
        return ""
    m = re.search(r"([+-]?\d{1,2})\s*C\s*(?:to|–|-)\s*([+-]?\d{1,2})\s*C", t, re.IGNORECASE)
    if m:
        l = int(m.group(1))
        r = int(m.group(2))
        left = f"+{l}" if l >= 0 else str(l)
        right = f"+{r}" if r >= 0 else str(r)
        return f"{left}C to {right}C"
    if "room temperature" in t.lower() or "ambient" in t.lower():
        return DEFAULT_STORAGE_TEMPERATURE_EN
    return t


def _temp_ru(temp_en: str) -> str:
    temp = _normalize_temp(temp_en) or DEFAULT_STORAGE_TEMPERATURE_EN
    temp = temp.replace(" to ", " до ")
    if temp.endswith(" ambient"):
        temp = temp.replace(" ambient", "")
    return temp


def _format_invoice_date(value: str) -> str:
    raw = _clean_space(value)
    if not raw:
        return ""

    candidates = [
        "%d-%b-%y",
        "%d-%b-%Y",
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d-%m-%y",
    ]

    normalized = raw.replace("/", "-").replace(".", "-")
    for fmt in candidates:
        try:
            dt = datetime.strptime(raw, fmt)
            return dt.strftime("%d.%m.%y")
        except ValueError:
            pass
        try:
            dt = datetime.strptime(normalized, fmt)
            return dt.strftime("%d.%m.%y")
        except ValueError:
            pass

    m = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", raw)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)[-2:]}"

    return raw


def _normalize_terms(value: str) -> str:
    text = _clean_space(value)
    if not text:
        return ""

    m = re.search(r"\b(CPT|FOB|CIF|EXW|DAP|DDP|FCA)\b", text, re.IGNORECASE)
    if not m:
        return text

    incoterm = m.group(1).upper()
    low = text.lower()
    has_air = bool(re.search(r"\b(by\s+air|air)\b", low))
    city = "MOSCOW" if "moscow" in low else ""

    result = incoterm
    if has_air:
        result += " BY AIR"
    if city:
        result += f" {city}"
    return result


def _cleanup_line(text: str) -> str:
    out = text
    out = re.sub(r"\(\s*([^()]*?)\s*,\s*\)", r"(\1)", out)
    out = out.replace(" ,", ",")
    out = re.sub(r"\s{2,}", " ", out)
    return out.strip()


def _replace_tokens(text: str, context: dict[str, str]) -> tuple[str, int, int]:
    replaced_known = 0
    removed_unknown = 0
    out = text

    for key, value in context.items():
        pattern = re.compile(r"{{\s*" + re.escape(key) + r"\s*}}")
        matches = list(pattern.finditer(out))
        if matches:
            replaced_known += len(matches)
            out = pattern.sub(value, out)

    unresolved = re.findall(r"{{\s*[A-Za-z0-9_]+\s*}}", out)
    if unresolved:
        removed_unknown += len(unresolved)
        out = re.sub(r"{{\s*[A-Za-z0-9_]+\s*}}", "", out)

    out = _cleanup_line(out)
    return out, replaced_known, removed_unknown


def fill_docx_template(template_bytes: bytes, context: dict[str, str]) -> tuple[bytes, int]:
    doc = Document(BytesIO(template_bytes))
    replaced_known = 0

    for paragraph in doc.paragraphs:
        new_text, known, unknown = _replace_tokens(paragraph.text or "", context)
        if known or unknown:
            paragraph.text = new_text
            replaced_known += known

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    new_text, known, unknown = _replace_tokens(paragraph.text or "", context)
                    if known or unknown:
                        paragraph.text = new_text
                        replaced_known += known

    output = BytesIO()
    doc.save(output)
    return output.getvalue(), replaced_known


def _build_context(
    data: ExtractedData,
    positions: list[Position],
    company_profile: dict[str, str] | None,
    temperature_en: str,
) -> dict[str, str]:
    company_profile = company_profile or {}

    exporter_company_name_en = company_profile.get("exporter_company_name_en") or data.exporter_name or ""
    exporter_company_name_ru = (
        company_profile.get("exporter_company_name_ru")
        or data.exporter_name_ru
        or exporter_company_name_en
    )
    exporter_company_address_en = company_profile.get("exporter_company_address_en") or data.exporter_address or ""

    storage_en = _normalize_temp(temperature_en) or DEFAULT_STORAGE_TEMPERATURE_EN
    storage_ru = company_profile.get("storage_temperature_ru") or _temp_ru(storage_en)

    context: dict[str, str] = {
        "INVOICE_NO": data.invoice_no or "",
        "INVOICE_DATE": _format_invoice_date(data.invoice_date),
        "TERMS_OF_DELIVERY": _normalize_terms(data.terms_of_delivery),
        "PERIOD_OF_VALIDITY": data.period_of_validity or "",
        "SPECIFICATION_DATE": _format_invoice_date(data.specification_date),
        "STORAGE_TEMPERATURE": storage_en,
        "STORAGE_TEMPERATURE_EN": storage_en,
        "STORAGE_TEMPERATURE_RU": storage_ru,
        "EXPORTER_COMPANY_NAME_EN": exporter_company_name_en,
        "EXPORTER_COMPANY_NAME_RU": exporter_company_name_ru,
        "EXPORTER_COMPANY_ADDRESS_EN": exporter_company_address_en,
        "EXPORTER_COMPANY_ADRESS_EN": exporter_company_address_en,
    }

    default_currency = data.currency or ""

    for index, position in enumerate(positions, start=1):
        qty = _qty_str(position.quantity)
        qty_en = f"{qty} un" if qty else ""
        qty_ru = f"{qty} шт" if qty else ""
        packing_en = position.packing_en or ""
        packing_ru = position.packing_ru or packing_en
        currency = position.currency or default_currency

        context[f"POSITION_{index}_NAME_EN"] = position.name_en or ""
        context[f"POSITION_{index}_NAME_RU"] = position.name_ru or position.name_en or ""
        context[f"POSITION_{index}_QUANTITY"] = qty
        context[f"POSITION_{index}_QUANTITY_EN"] = qty_en
        context[f"POSITION_{index}_QUANTITY_RU"] = qty_ru
        context[f"POSITION_{index}_PACKING"] = packing_en
        context[f"POSITION_{index}_PACKING_EN"] = packing_en
        context[f"POSITION_{index}_PACKING_RU"] = packing_ru
        context[f"POSITION_{index}_UNIT_PRICE"] = _money_with_currency(position.unit_price, currency)
        context[f"POSITION_{index}_TOTAL_PRICE"] = _money_with_currency(position.total_price, currency)
        context[f"POSITION_{index}_CURRENCY"] = currency

    return context


def _build_default_price_doc(data: ExtractedData, positions: list[Position], company_profile: dict[str, str] | None) -> bytes:
    _ = company_profile
    doc = Document()
    _set_default_font(doc)

    p = doc.add_paragraph("On the company letterhead")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title = doc.add_paragraph("PRICE LIST")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    date_paragraph = doc.add_paragraph(_format_invoice_date(data.invoice_date) or "")
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for idx, position in enumerate(positions):
        if idx > 0:
            doc.add_paragraph()
        qty = _qty_str(position.quantity)
        currency = position.currency or data.currency or ""
        doc.add_paragraph(f"PRODUCT NAME: {position.name_en or '-'}")
        doc.add_paragraph(f"QUANITTY: {qty or '-'}")
        doc.add_paragraph(f"PACKING: {position.packing_en or '-'}")
        doc.add_paragraph(f"UNIT PRICE: {_money_with_currency(position.unit_price, currency) or '-'}")
        doc.add_paragraph(f"TOTAL AMOUNT: {_money_with_currency(position.total_price, currency) or '-'}")

    doc.add_paragraph()
    doc.add_paragraph(f"TERMS OF DELIVERY: {_normalize_terms(data.terms_of_delivery) or '-'}")
    doc.add_paragraph(f"PERIOD OF VALIDITY: {data.period_of_validity or '-'}")
    doc.add_paragraph()
    doc.add_paragraph("stamp / signature")

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def generate_price_list_doc(
    data: ExtractedData,
    company_info: str,
    template_bytes: bytes | None = None,
    company_profile: dict[str, str] | None = None,
) -> bytes:
    _ = company_info
    positions = data.positions or [Position()]
    temperature = positions[0].storage_temperature or data.storage_temperature or DEFAULT_STORAGE_TEMPERATURE_EN
    context = _build_context(data, positions, company_profile, temperature)

    if template_bytes:
        templated, replaced = fill_docx_template(template_bytes, context)
        if replaced > 0:
            return templated

    return _build_default_price_doc(data, positions, company_profile)


def _default_label_doc(
    data: ExtractedData,
    positions: list[Position],
    temperature_en: str,
    company_profile: dict[str, str] | None,
) -> bytes:
    context = _build_context(data, positions, company_profile, temperature_en)

    doc = Document()
    _set_default_font(doc)

    header = doc.add_paragraph("Наименование товара / Product name - Quantity / Кол-во:")
    header.runs[0].bold = True

    for pos_index, _position in enumerate(positions, start=1):
        line = (
            f"{context.get(f'POSITION_{pos_index}_NAME_EN', '')} / {context.get(f'POSITION_{pos_index}_NAME_RU', '')} "
            f"{context.get(f'POSITION_{pos_index}_QUANTITY_EN', '')} "
            f"({context.get(f'POSITION_{pos_index}_PACKING_EN', '')}) / "
            f"{context.get(f'POSITION_{pos_index}_QUANTITY_RU', '')} "
            f"({context.get(f'POSITION_{pos_index}_PACKING_RU', '')})"
        ).strip()
        doc.add_paragraph(_cleanup_line(line))

    doc.add_paragraph(
        f"Shipping Conditions: Require temperature-controlled shipping must be kept between ({context['STORAGE_TEMPERATURE_EN']})"
    )
    doc.add_paragraph(
        f"Условия транспортировки: требуется контролируемая температура при транспортировке ({context['STORAGE_TEMPERATURE_RU']})"
    )
    doc.add_paragraph(
        f"Shipper/Отправитель: \"{context['EXPORTER_COMPANY_NAME_EN']}\" / «{context['EXPORTER_COMPANY_NAME_RU']}»"
    )
    doc.add_paragraph(f"Contact information / Контактная информация: {context['EXPORTER_COMPANY_ADRESS_EN']}")

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def _temperature_slug(temp: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", temp).strip("_")
    return slug[:60] or "ambient"


def _group_positions_by_temperature(data: ExtractedData) -> OrderedDict[str, list[Position]]:
    grouped: OrderedDict[str, list[Position]] = OrderedDict()
    positions = data.positions or [Position()]

    for position in positions:
        temp = _normalize_temp(position.storage_temperature) or _normalize_temp(data.storage_temperature)
        if not temp:
            temp = DEFAULT_STORAGE_TEMPERATURE_EN
        if temp not in grouped:
            grouped[temp] = []
        grouped[temp].append(position)

    return grouped


def generate_label_docs_by_temperature(
    data: ExtractedData,
    company_info: str,
    template_bytes: bytes | None = None,
    company_profile: dict[str, str] | None = None,
) -> OrderedDict[str, bytes]:
    _ = company_info
    grouped = _group_positions_by_temperature(data)
    docs: OrderedDict[str, bytes] = OrderedDict()
    many = len(grouped) > 1

    for temperature, positions in grouped.items():
        context = _build_context(data, positions, company_profile, temperature)
        label_bytes: bytes | None = None

        if template_bytes:
            templated, replaced = fill_docx_template(template_bytes, context)
            if replaced > 0:
                label_bytes = templated

        if label_bytes is None:
            label_bytes = _default_label_doc(data, positions, temperature, company_profile)

        filename = f"Label_{_temperature_slug(temperature)}.docx" if many else "Label.docx"
        docs[filename] = label_bytes

    return docs


def generate_label_doc(
    data: ExtractedData,
    company_info: str,
    template_bytes: bytes | None = None,
    company_profile: dict[str, str] | None = None,
) -> bytes:
    docs = generate_label_docs_by_temperature(
        data=data,
        company_info=company_info,
        template_bytes=template_bytes,
        company_profile=company_profile,
    )
    return next(iter(docs.values()))
