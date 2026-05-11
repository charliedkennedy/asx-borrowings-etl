from typing import List, Tuple
import fitz


def extract_pdf_text_by_page(pdf_path: str) -> Tuple[List[str], str]:
    try:
        doc = fitz.open(pdf_path)
        pages = []
        total_chars = 0
        for p in doc:
            t = p.get_text("text") or ""
            pages.append(t)
            total_chars += len(t.strip())
        quality = "ok" if total_chars >= 1000 else "poor"
        return pages, quality
    except Exception:
        return [], "failed"
