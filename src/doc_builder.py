from __future__ import annotations

from io import BytesIO
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from .models import ExtractedData, Position


def _set_default_font(document: Document, font_name: str = "Times New Roman", size_pt: int = 12) -> None:
    style = document.styles["Normal"]
    style.font.name = font_name
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    style.font.size = Pt(size_pt)


def _add_key_value(document: Document, key: str, value: str) -> None:
    paragraph = document.add_paragraph()
    run_k = paragraph.add_run(f"{key}: ")
    run_k.bold = True
    paragraph.add_run(value or "-")


def _price_str(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value:,.2f}"


def _position_context(positions: list[Position]) -> str:
    rows: list[str] = []
    for idx, pos in enumerate(positions, start=1):
        rows.append(
            " | ".join(
                [
                    str(idx),
                    pos.code or "-",
                    pos.name_en or "-",
                    str(pos.quantity if pos.quantity is not None else "-"),
                    pos.packing_en or "-",
                    _price_str(pos.unit_price),
                    _price_str(pos.total_price),
                    pos.currency or "-",
                ]
            )
        )
    return "\n".join(rows)


def _build_template_context(data: ExtractedData, company_info: str) -> dict[str, str]:
    first = data.positions[0] if data.positions else Position()
    return {
        "INVOICE_NO": data.invoice_no,
        "INVOICE_DATE": data.invoice_date,
        "BUYER_NAME": data.buyer_name,
        "BUYER_ADDRESS": data.buyer_address,
        "EXPORTER_NAME": data.exporter_name,
        "EXPORTER_ADDRESS": data.exporter_address,
        "TERMS_OF_DELIVERY": data.terms_of_delivery,
        "PERIOD_OF_VALIDITY": data.period_of_validity,
        "SPECIFICATION_DATE": data.specification_date,
        "STORAGE_TEMPERATURE": data.storage_temperature,
        "COMPANY_INFO": company_info,
        "POSITIONS_TABLE": _position_context(data.positions),
        "POSITION_1_NAME_EN": first.name_en,
        "POSITION_1_NAME_RU": first.name_ru,
        "POSITION_1_QUANTITY": str(first.quantity or ""),
        "POSITION_1_PACKING_EN": first.packing_en,
        "POSITION_1_PACKING_RU": first.packing_ru,
        "POSITION_1_PRICE": _price_str(first.unit_price),
        "POSITION_1_TOTAL": _price_str(first.total_price),
        "POSITION_1_CURRENCY": first.currency or data.currency,
    }


def _replace_placeholders_in_paragraph(paragraph, context: dict[str, str]) -> None:
    # python-docx run-level replacement is fragile; replace full paragraph text to make simple templates usable.
    text = paragraph.text
    changed = False
    for k, v in context.items():
        key = f"{{{{{k}}}}}"
        if key in text:
            text = text.replace(key, v or "")
            changed = True
    if changed:
        paragraph.text = text


def fill_docx_template(template_bytes: bytes, context: dict[str, str]) -> bytes:
    doc = Document(BytesIO(template_bytes))

    for p in doc.paragraphs:
        _replace_placeholders_in_paragraph(p, context)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_placeholders_in_paragraph(p, context)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def generate_price_list_doc(
    data: ExtractedData,
    company_info: str,
    template_bytes: bytes | None = None,
) -> bytes:
    context = _build_template_context(data, company_info)
    if template_bytes:
        return fill_docx_template(template_bytes, context)

    doc = Document()
    _set_default_font(doc)

    p = doc.add_paragraph("On the company letterhead")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title = doc.add_paragraph("PRICE LIST")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    _add_key_value(doc, "Date", data.specification_date or data.invoice_date)
    _add_key_value(doc, "Invoice No", data.invoice_no)
    _add_key_value(doc, "Buyer", data.buyer_name)
    _add_key_value(doc, "Buyer Address", data.buyer_address)
    _add_key_value(doc, "Exporter", data.exporter_name)
    _add_key_value(doc, "Exporter Address", data.exporter_address)

    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=8)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    headers = ["#", "Code", "Product Name (EN)", "Qty", "Packing", "Unit Price", "Total", "Currency"]
    for idx, h in enumerate(headers):
        hdr[idx].text = h

    for i, position in enumerate(data.positions, start=1):
        row = table.add_row().cells
        row[0].text = str(i)
        row[1].text = position.code or "-"
        row[2].text = position.name_en or "-"
        row[3].text = str(position.quantity if position.quantity is not None else "-")
        row[4].text = position.packing_en or "-"
        row[5].text = _price_str(position.unit_price)
        row[6].text = _price_str(position.total_price)
        row[7].text = position.currency or data.currency or "-"

    doc.add_paragraph()
    _add_key_value(doc, "Terms of Delivery", data.terms_of_delivery)
    _add_key_value(doc, "Period of Validity", data.period_of_validity)
    _add_key_value(doc, "Date of Specification", data.specification_date)

    if company_info.strip():
        doc.add_paragraph()
        _add_key_value(doc, "Company Info", company_info.strip())

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def generate_label_doc(
    data: ExtractedData,
    company_info: str,
    template_bytes: bytes | None = None,
) -> bytes:
    context = _build_template_context(data, company_info)
    if template_bytes:
        return fill_docx_template(template_bytes, context)

    doc = Document()
    _set_default_font(doc)

    if not data.positions:
        data.positions = [Position()]

    for idx, position in enumerate(data.positions):
        table = doc.add_table(rows=3, cols=1)
        table.style = "Table Grid"

        header = table.rows[0].cells[0].paragraphs[0]
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hrun = header.add_run("Наименование товара / Product name - Quantity / Кол-во")
        hrun.bold = True

        middle = table.rows[1].cells[0].paragraphs[0]
        middle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ru_name = position.name_ru or "[RU name]"
        ru_pack = position.packing_ru or position.packing_en or ""
        qty = f"{position.quantity:g}" if position.quantity is not None else ""
        middle.add_run(
            f"{position.name_en or '[EN name]'} / {ru_name} - {qty} {position.packing_en} / {ru_pack}".strip()
        )

        bottom = table.rows[2].cells[0].paragraphs[0]
        bottom.alignment = WD_ALIGN_PARAGRAPH.CENTER
        bottom.add_run(f"Storage: {data.storage_temperature or '[Not extracted]'}")

        if company_info.strip():
            doc.add_paragraph(company_info.strip())

        if idx < len(data.positions) - 1:
            doc.add_page_break()

    output = BytesIO()
    doc.save(output)
    return output.getvalue()
