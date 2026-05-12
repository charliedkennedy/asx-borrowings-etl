from typing import List, Tuple


def _extract_with_pymupdf(pdf_path: str) -> Tuple[List[str], str]:
    import fitz

    doc = fitz.open(pdf_path)
    pages: List[str] = []
    total_chars = 0
    for page in doc:
        text = page.get_text("text") or ""
        pages.append(text)
        total_chars += len(text.strip())
    if not pages:
        return [], "failed"
    if total_chars >= 1000:
        return pages, "ok"
    if total_chars >= 200:
        return pages, "poor"
    return pages, "very poor"


def _extract_with_pypdf(pdf_path: str) -> Tuple[List[str], str]:
    from pypdf import PdfReader

    reader = PdfReader(pdf_path)
    pages: List[str] = []
    total_chars = 0
    for page in reader.pages:
        text = page.extract_text() or ""
        pages.append(text)
        total_chars += len(text.strip())
    if not pages:
        return [], "failed"
    if total_chars >= 1000:
        return pages, "ok"
    if total_chars >= 200:
        return pages, "poor"
    return pages, "very poor"


def _extract_with_pdfplumber(pdf_path: str) -> Tuple[List[str], str]:
    import pdfplumber

    pages: List[str] = []
    total_chars = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            pages.append(text)
            total_chars += len(text.strip())
    if not pages:
        return [], "failed"
    if total_chars >= 1000:
        return pages, "ok"
    if total_chars >= 200:
        return pages, "poor"
    return pages, "very poor"


def extract_pdf_text_by_page(pdf_path: str) -> Tuple[List[str], str]:
    """Extract text by page. Uses installed libraries in priority order.

    Returns (pages, quality). Quality can be:
    - ok
    - poor
    - very poor
    - failed
    """
    errors = []
    for extractor in (_extract_with_pymupdf, _extract_with_pypdf, _extract_with_pdfplumber):
        try:
            pages, quality = extractor(pdf_path)
            if pages and quality in {"ok", "poor"}:
                return pages, quality
            if pages:
                # Keep the best available text even if weak; caller will flag manual review.
                return pages, quality
        except Exception as exc:
            errors.append(f"{extractor.__name__}: {exc}")
    return ["\n".join(errors)], "failed"
