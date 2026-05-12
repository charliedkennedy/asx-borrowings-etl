import re
from pathlib import Path
from typing import Optional, Tuple
from urllib.parse import parse_qs, urljoin, urlparse

import requests


PDF_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
    "Accept": "application/pdf,application/octet-stream,text/html,application/xhtml+xml,*/*;q=0.8",
    "Accept-Language": "en-AU,en;q=0.9",
    "Referer": "https://www.asx.com.au/markets/trade-our-cash-market/announcements",
}


def _is_pdf_response(data: bytes, content_type: str) -> bool:
    return data.startswith(b"%PDF") or "pdf" in (content_type or "").lower()


def _write_pdf(path: Path, data: bytes) -> Tuple[str, Optional[str]]:
    if not data.startswith(b"%PDF"):
        # Some servers return application/pdf with leading whitespace/BOM. Trim only before %PDF.
        idx = data.find(b"%PDF")
        if idx >= 0:
            data = data[idx:]
    path.write_bytes(data)
    return str(path), None


def _extract_pdf_links(html: str, base_url: str) -> list[str]:
    links = []
    for m in re.finditer(r"https?://[^\s'\"<>]+\.pdf(?:\?[^\s'\"<>]*)?", html, flags=re.I):
        links.append(m.group(0))
    for m in re.finditer(r"(?:href|src)=[\'\"]([^\'\"]+\.pdf(?:\?[^\'\"]*)?)[\'\"]", html, flags=re.I):
        links.append(urljoin(base_url, m.group(1)))
    # ASX direct PDF links often use this path even if not ending in .pdf in the HTML.
    for m in re.finditer(r"https?://announcements\.asx\.com\.au/asxpdf/[^\s'\"<>]+", html, flags=re.I):
        links.append(m.group(0))
    seen = set()
    out = []
    for link in links:
        if link not in seen:
            seen.add(link)
            out.append(link)
    return out


def _asx_display_variants(url: str) -> list[str]:
    """Return possible ASX display URL variants.

    The ASX historical announcement pages expose displayAnnouncement links. Some
    environments are fussy about parameter casing / redirects, so try both forms.
    """
    variants = [url]
    parsed = urlparse(url)
    if "displayannouncement.do" not in parsed.path.lower():
        return variants
    qs = parse_qs(parsed.query)
    ids = (qs.get("idsId") or qs.get("idsID") or qs.get("idsid") or [""])[0]
    if ids:
        variants.extend([
            f"https://www.asx.com.au/asx/v2/statistics/displayAnnouncement.do?display=pdf&idsId={ids}",
            f"https://www.asx.com.au/asx/v2/statistics/displayAnnouncement.do?display=pdf&idsID={ids}",
            f"https://www.asx.com.au/asxpdf/{ids}.pdf",
        ])
    # De-duplicate while preserving order.
    seen = set()
    return [v for v in variants if not (v in seen or seen.add(v))]


def _try_get(session: requests.Session, url: str, timeout: int):
    return session.get(url, headers=PDF_HEADERS, timeout=timeout, allow_redirects=True)


def download_pdf(url: str, ticker: str, pdf_cache: Path, timeout: int = 30):
    pdf_cache.mkdir(parents=True, exist_ok=True)
    safe_ticker = (ticker or "company").replace("/", "_")
    out = pdf_cache / f"{safe_ticker}.pdf"
    if out.exists() and out.stat().st_size > 0:
        return str(out), None

    session = requests.Session()
    errors = []
    urls_to_try = _asx_display_variants(url)

    for candidate_url in urls_to_try:
        try:
            response = _try_get(session, candidate_url, timeout)
            data = response.content or b""
            content_type = response.headers.get("Content-Type", "")
            final_url = response.url
            if response.status_code >= 400:
                errors.append(f"{candidate_url} status={response.status_code}")
                continue
            if _is_pdf_response(data, content_type):
                return _write_pdf(out, data)

            # If ASX/company returns a landing page, follow embedded PDF links.
            text = data[:250000].decode(response.encoding or "utf-8", errors="ignore")
            for pdf_link in _extract_pdf_links(text, final_url):
                try:
                    pdf_response = _try_get(session, pdf_link, timeout)
                    pdf_data = pdf_response.content or b""
                    pdf_ctype = pdf_response.headers.get("Content-Type", "")
                    if pdf_response.status_code < 400 and _is_pdf_response(pdf_data, pdf_ctype):
                        return _write_pdf(out, pdf_data)
                    errors.append(f"embedded PDF failed {pdf_link} status={pdf_response.status_code} ctype={pdf_ctype}")
                except Exception as exc:
                    errors.append(f"embedded PDF failed {pdf_link}: {exc}")

            sample = re.sub(r"\s+", " ", text[:250]).strip()
            errors.append(f"{candidate_url} not PDF status={response.status_code} ctype={content_type} sample={sample}")
        except Exception as exc:
            errors.append(f"{candidate_url} error={exc}")

    return "", " | ".join(errors[:5]) or "URL did not return PDF content"
