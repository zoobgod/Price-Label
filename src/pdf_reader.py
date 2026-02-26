from __future__ import annotations

import re
import shutil
from io import BytesIO

import fitz
import pytesseract
from PIL import Image, ImageOps
from pypdf import PdfReader

from .models import ExtractionMeta, TextBundle


def _extract_text_native(pdf_bytes: bytes) -> list[str]:
    reader = PdfReader(BytesIO(pdf_bytes))
    pages: list[str] = []
    for page in reader.pages:
        pages.append((page.extract_text() or "").strip())
    return pages


def _render_page_image(doc: fitz.Document, page_index: int, dpi: int = 300) -> Image.Image:
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = doc.load_page(page_index).get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    gray = ImageOps.grayscale(img)
    return ImageOps.autocontrast(gray)


def _extract_text_ocr_page(doc: fitz.Document, page_index: int, languages: str = "eng+rus") -> str:
    image = _render_page_image(doc, page_index)
    config = "--oem 3 --psm 6"
    try:
        return pytesseract.image_to_string(image, lang=languages, config=config).strip()
    except pytesseract.TesseractError:
        return pytesseract.image_to_string(image, lang="eng", config=config).strip()


def _resolve_ocr_languages(requested: str) -> str:
    requested_parts = [p.strip() for p in requested.split("+") if p.strip()]
    if not requested_parts:
        return "eng"
    try:
        available = set(pytesseract.get_languages(config=""))
    except Exception:
        return "eng"
    allowed = [lang for lang in requested_parts if lang in available]
    if allowed:
        return "+".join(allowed)
    return "eng"


def _quality_score(text: str) -> int:
    if not text:
        return 0
    cleaned = re.sub(r"\s+", " ", text).strip()
    score = len(cleaned)
    keywords = [
        "invoice",
        "consignee",
        "description of goods",
        "quantity",
        "terms of delivery",
        "specification",
        "storage",
        "temperature",
    ]
    lowered = cleaned.lower()
    score += sum(80 for kw in keywords if kw in lowered)
    if re.search(r"\b[A-Z]{2,}/[A-Z]/\d{2}-\d{2}/\d+\b", cleaned):
        score += 120
    if re.search(r"\d{1,3}(?:,\d{2,3})+(?:\.\d+)?", cleaned):
        score += 40
    return score


def _normalize_ocr_noise(text: str) -> str:
    text = text.replace("’", "'")
    # Common OCR confusion around rupee and delimiters.
    text = text.replace("INR)", "(In INR)")
    text = re.sub(r"\s{2,}", " ", text)
    return text


def extract_pdf_text(
    pdf_bytes: bytes,
    force_ocr: bool = False,
    min_native_chars: int = 60,
    ocr_languages: str = "eng+rus",
    prefer_ocr: bool = False,
) -> TextBundle:
    """Extract text from PDF with OCR fallback for scanned/low-text pages."""
    tesseract_available = shutil.which("tesseract") is not None

    native_pages = _extract_text_native(pdf_bytes)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    ocr_languages = _resolve_ocr_languages(ocr_languages)

    output_pages: list[str] = []
    per_page_source: list[str] = []
    pages_ocrd = 0

    for idx, native_text in enumerate(native_pages):
        should_ocr = force_ocr or prefer_ocr or len(native_text) < min_native_chars
        if not (should_ocr and tesseract_available):
            output_pages.append(native_text)
            per_page_source.append("native")
            continue

        ocr_text = _normalize_ocr_noise(_extract_text_ocr_page(doc, idx, languages=ocr_languages))
        native_score = _quality_score(native_text)
        ocr_score = _quality_score(ocr_text)

        # Prefer OCR when explicitly requested or when OCR text appears structurally better.
        if force_ocr or (prefer_ocr and ocr_score >= native_score * 0.8) or (ocr_score > native_score):
            output_pages.append(ocr_text)
            per_page_source.append("ocr")
            pages_ocrd += 1
        else:
            output_pages.append(native_text)
            per_page_source.append("native")

    text = "\n\n".join(output_pages).strip()
    meta = ExtractionMeta(
        tesseract_available=tesseract_available,
        pages_total=len(output_pages),
        pages_ocrd=pages_ocrd,
        per_page_source=per_page_source,
    )
    return TextBundle(text=text, meta=meta)
