from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class Position:
    code: str = ""
    name_en: str = ""
    name_ru: str = ""
    quantity: float | None = None
    packing_en: str = ""
    packing_ru: str = ""
    unit_price: float | None = None
    total_price: float | None = None
    currency: str = ""
    storage_temperature: str = ""


@dataclass
class ExtractionMeta:
    tesseract_available: bool = False
    pages_total: int = 0
    pages_ocrd: int = 0
    per_page_source: list[str] = field(default_factory=list)


@dataclass
class TextBundle:
    text: str
    meta: ExtractionMeta


@dataclass
class ExtractedData:
    invoice_no: str = ""
    invoice_date: str = ""
    buyer_name: str = ""
    buyer_address: str = ""
    exporter_name: str = ""
    exporter_name_ru: str = ""
    exporter_address: str = ""
    terms_of_delivery: str = ""
    period_of_validity: str = ""
    specification_date: str = ""
    storage_temperature: str = ""
    positions: list[Position] = field(default_factory=list)
    currency: str = ""
    raw: dict[str, Any] = field(default_factory=dict)
