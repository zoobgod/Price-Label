from __future__ import annotations

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
    return pytesseract.image_to_string(image, lang=languages, config=config).strip()


def extract_pdf_text(
    pdf_bytes: bytes,
    force_ocr: bool = False,
    min_native_chars: int = 60,
    ocr_languages: str = "eng+rus",
) -> TextBundle:
    """Extract text from PDF with OCR fallback for scanned/low-text pages."""
    tesseract_available = shutil.which("tesseract") is not None

    native_pages = _extract_text_native(pdf_bytes)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    output_pages: list[str] = []
    per_page_source: list[str] = []
    pages_ocrd = 0

    for idx, native_text in enumerate(native_pages):
        should_ocr = force_ocr or len(native_text) < min_native_chars
        if should_ocr and tesseract_available:
            ocr_text = _extract_text_ocr_page(doc, idx, languages=ocr_languages)
            if len(ocr_text) >= len(native_text):
                output_pages.append(ocr_text)
                per_page_source.append("ocr")
                pages_ocrd += 1
                continue

        output_pages.append(native_text)
        per_page_source.append("native")

    if force_ocr and tesseract_available:
        for idx in range(len(native_pages)):
            if per_page_source[idx] == "native":
                ocr_text = _extract_text_ocr_page(doc, idx, languages=ocr_languages)
                if len(ocr_text) > len(output_pages[idx]):
                    output_pages[idx] = ocr_text
                    per_page_source[idx] = "ocr"
                    pages_ocrd += 1

    text = "\n\n".join(output_pages).strip()
    meta = ExtractionMeta(
        tesseract_available=tesseract_available,
        pages_total=len(output_pages),
        pages_ocrd=pages_ocrd,
        per_page_source=per_page_source,
    )
    return TextBundle(text=text, meta=meta)
