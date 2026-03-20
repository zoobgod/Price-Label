from __future__ import annotations

import base64
from io import BytesIO

import fitz
from PIL import Image


def pdf_to_page_images(pdf_bytes: bytes, dpi: int = 150, max_pages: int = 20) -> list[bytes]:
    """Render each PDF page as a PNG image (bytes)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    images: list[bytes] = []
    for idx in range(min(len(doc), max_pages)):
        pix = doc.load_page(idx).get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        buf = BytesIO()
        img.save(buf, format="PNG", optimize=True)
        images.append(buf.getvalue())
    return images


def pdf_to_page_images_b64(pdf_bytes: bytes, dpi: int = 150, max_pages: int = 20) -> list[str]:
    """Render each PDF page as a base64-encoded PNG string."""
    return [base64.b64encode(img).decode() for img in pdf_to_page_images(pdf_bytes, dpi, max_pages)]


def pdf_to_native_text(pdf_bytes: bytes) -> str:
    """Extract machine-readable text from PDF using PyMuPDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages: list[str] = []
    for page in doc:
        pages.append(page.get_text().strip())
    return "\n\n".join(pages)
