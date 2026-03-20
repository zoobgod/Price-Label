"""LLM-powered extraction using Claude Vision API with structured tool_use output."""

from __future__ import annotations

import base64

from anthropic import Anthropic

MODEL = "claude-sonnet-4-20250514"

# ---------------------------------------------------------------------------
# Tool schemas – force Claude to return structured JSON via tool_use
# ---------------------------------------------------------------------------

PI_TOOL = {
    "name": "submit_pi_data",
    "description": "Submit all extracted data from a Proforma Invoice / Commercial Invoice.",
    "input_schema": {
        "type": "object",
        "properties": {
            "invoice_no": {
                "type": "string",
                "description": "Invoice number exactly as printed (e.g. BIPL/PI/25-26/026).",
            },
            "invoice_date": {
                "type": "string",
                "description": "Invoice date normalised to DD.MM.YY (e.g. 04.09.25).",
            },
            "buyer_name": {"type": "string"},
            "buyer_address": {"type": "string"},
            "exporter_name": {
                "type": "string",
                "description": "Name of the exporting / selling company.",
            },
            "exporter_address": {
                "type": "string",
                "description": "Full address of the exporter, on one line.",
            },
            "currency": {
                "type": "string",
                "description": "ISO-4217 currency code found on the invoice (INR, USD, EUR …).",
            },
            "terms_of_delivery": {
                "type": "string",
                "description": (
                    "Delivery incoterm + destination ONLY. "
                    "Example: 'CPT BY AIR MOSCOW'. "
                    "Do NOT include 'INCOTERMS 2020' or any version / year text."
                ),
            },
            "positions": {
                "type": "array",
                "description": "Every product line-item on the invoice.",
                "items": {
                    "type": "object",
                    "properties": {
                        "code": {
                            "type": "string",
                            "description": "Product / catalogue / part code.",
                        },
                        "name_en": {
                            "type": "string",
                            "description": "Product name in English.",
                        },
                        "quantity": {"type": "number"},
                        "packing": {
                            "type": "string",
                            "description": "Packing description, e.g. '1 x 50MG', '01 x 100mg'.",
                        },
                        "unit_price": {
                            "type": "number",
                            "description": "Unit price as a plain number (no currency symbol).",
                        },
                        "total_price": {
                            "type": "number",
                            "description": "Total price as a plain number.",
                        },
                    },
                    "required": ["name_en"],
                },
            },
        },
        "required": ["positions"],
    },
}

MSDS_TOOL = {
    "name": "submit_msds_data",
    "description": "Submit extracted storage / shipping temperature from an MSDS.",
    "input_schema": {
        "type": "object",
        "properties": {
            "storage_temperature": {
                "type": "string",
                "description": (
                    "Storage or shipping temperature range. "
                    "Use the format '+2C to +8C' or '+15C to +25C ambient'. "
                    "If 'room temperature' / 'ambient' without numbers, return '+15C to +25C ambient'. "
                    "If nothing found, return empty string."
                ),
            },
        },
        "required": ["storage_temperature"],
    },
}

SPEC_TOOL = {
    "name": "submit_spec_data",
    "description": "Submit extracted data from a Specification / Contract document.",
    "input_schema": {
        "type": "object",
        "properties": {
            "terms_of_delivery": {
                "type": "string",
                "description": "Delivery incoterm + destination ONLY (e.g. 'CPT BY AIR MOSCOW').",
            },
            "period_of_validity": {
                "type": "string",
                "description": "Validity period or shipment timeframe, e.g. 'September - November 2025'.",
            },
            "specification_date": {
                "type": "string",
                "description": "Specification date normalised to DD.MM.YY.",
            },
        },
    },
}

# ---------------------------------------------------------------------------
# Prompts
# ---------------------------------------------------------------------------

PI_PROMPT = """\
You are an expert document-data extractor. The images show a Proforma Invoice \
(or Commercial Invoice) for pharmaceutical / pharmacopoeia products, often \
from Indian exporters.

Extract every piece of data you can find and submit it via the tool.

Key rules:
* invoice_date → always DD.MM.YY  (e.g. 26.02.26, 04.09.25).
* terms_of_delivery → ONLY the incoterm keyword + transport mode + city. \
  Example: "CPT BY AIR MOSCOW".  Never include "INCOTERMS 2020" or similar.
* positions → one entry per product row in the goods table.  \
  Include code, English name, quantity, packing, unit_price, total_price.  \
  Prices are plain numbers – no currency symbols.
* currency → single ISO code for the whole invoice.
* If the document is a scan / image, do your best to read every field.
"""

MSDS_PROMPT = """\
You are extracting the storage / shipping temperature from a Material Safety \
Data Sheet (MSDS) or Certificate of Analysis for a pharmaceutical product.

Look for any of: storage conditions, recommended temperature, shipping \
temperature, "store at …", "keep between …".

Rules:
* Format as "+2C to +8C" or "+15C to +25C ambient".
* "Room temperature" or "ambient" without specific numbers → "+15C to +25C ambient".
* If no temperature is found, return an empty string.

Submit via the tool.
"""

SPEC_PROMPT = """\
You are extracting data from a pharmaceutical Specification or Contract document.

Extract:
* terms_of_delivery – incoterm + destination only (e.g. "CPT BY AIR MOSCOW").
* period_of_validity – the validity / shipment period (e.g. "September - November 2025").
* specification_date – the document date in DD.MM.YY format.

Submit via the tool.
"""

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _image_content_blocks(page_images: list[bytes]) -> list[dict]:
    blocks: list[dict] = []
    for img_bytes in page_images:
        blocks.append(
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": base64.b64encode(img_bytes).decode(),
                },
            }
        )
    return blocks


def _call_with_tool(
    client: Anthropic,
    page_images: list[bytes],
    native_text: str,
    prompt: str,
    tool: dict,
) -> dict:
    """Send page images + optional native text to Claude and force a tool_use response."""
    content = _image_content_blocks(page_images)

    # Append native text as supplementary context (helps with numbers / codes)
    trimmed = (native_text or "").strip()
    if trimmed:
        content.append(
            {
                "type": "text",
                "text": (
                    "Machine-readable text extracted from the same PDF "
                    "(may be incomplete, mis-ordered, or empty for scanned pages):\n\n"
                    + trimmed[:10_000]
                ),
            }
        )

    content.append({"type": "text", "text": prompt})

    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        tools=[tool],
        tool_choice={"type": "tool", "name": tool["name"]},
        messages=[{"role": "user", "content": content}],
    )

    for block in response.content:
        if block.type == "tool_use":
            return block.input  # type: ignore[return-value]
    return {}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def extract_pi(client: Anthropic, page_images: list[bytes], native_text: str) -> dict:
    return _call_with_tool(client, page_images, native_text, PI_PROMPT, PI_TOOL)


def extract_msds(client: Anthropic, page_images: list[bytes], native_text: str) -> dict:
    return _call_with_tool(client, page_images, native_text, MSDS_PROMPT, MSDS_TOOL)


def extract_spec(client: Anthropic, page_images: list[bytes], native_text: str) -> dict:
    return _call_with_tool(client, page_images, native_text, SPEC_PROMPT, SPEC_TOOL)
