import re
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, List, Optional, Tuple
from urllib.parse import parse_qs, quote, unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from .schemas import SourceRecord

REPORT_KEYWORDS = [
    "annual report",
    "appendix 4e",
    "appendix 4d",
    "half year",
    "half-year",
    "half yearly",
    "half-yearly",
    "financial report",
    "report and accounts",
    "preliminary final report",
    "full year report",
]

LOW_VALUE_KEYWORDS = [
    "presentation",
    "media release",
    "transcript",
    "notification",
    "notice of meeting",
    "director",
    "substantial holder",
    "buy-back",
    "dividend",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; asx-borrowings-etl/1.0; +https://github.com/)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


@dataclass
class Candidate:
    title: str
    url: str
    source_type: str
    source_page: str = ""
    score: int = 0


def _asx_code(ticker: str) -> str:
    return (ticker or "").upper().replace(".AX", "").strip()


def _decode_search_href(href: str) -> str:
    """Decode DuckDuckGo redirect links into the real target URL."""
    if not href:
        return ""
    if href.startswith("//"):
        href = "https:" + href
    parsed = urlparse(href)
    qs = parse_qs(parsed.query)
    if "uddg" in qs and qs["uddg"]:
        return unquote(qs["uddg"][0])
    return href


def _is_official(url: str) -> bool:
    u = (url or "").lower()
    if not u.startswith(("http://", "https://")):
        return False
    if "asx.com.au" in u or "announcements.asx.com.au" in u:
        return True
    # Company investor-centre PDFs/pages are usually on the issuer's own domain.
    # Exclude obvious third-party market data / news aggregators.
    blocked = [
        "marketindex.com.au",
        "investsmart.com.au",
        "morningstar",
        "reuters",
        "bloomberg",
        "afr.com",
        "themarketonline",
    ]
    return not any(b in u for b in blocked) and ".com" in urlparse(u).netloc


def _looks_like_pdf_or_announcement(url: str) -> bool:
    u = (url or "").lower()
    return u.endswith(".pdf") or "/asxpdf/" in u or "displayannouncement" in u or "announcements.asx.com.au" in u


def _looks_report(text: str, url: str = "") -> bool:
    combined = f"{text or ''} {url or ''}".lower()
    if not any(k in combined for k in REPORT_KEYWORDS):
        return False
    if "presentation" in combined and not ("appendix 4d" in combined or "appendix 4e" in combined or "financial report" in combined):
        return False
    return True


def _extract_report_date(title: str, url: str = "") -> str:
    combined = f"{title or ''} {url or ''}"
    # Prefer explicit day/month/year dates where available.
    patterns = [
        r"(\d{1,2})[\-/\.](\d{1,2})[\-/\.](20\d{2})",
        r"(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+(20\d{2})",
    ]
    month_lookup = {m.lower(): i for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}
    month_lookup["sept"] = 9
    for p in patterns:
        m = re.search(p, combined, flags=re.I)
        if not m:
            continue
        if m.group(2).isdigit():
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            day, month, year = int(m.group(1)), month_lookup[m.group(2).lower()[:4] if m.group(2).lower().startswith("sept") else m.group(2).lower()[:3]], int(m.group(3))
        try:
            return datetime(year, month, day).date().isoformat()
        except ValueError:
            pass
    # Otherwise capture financial/reporting year only as reporting_period, not report date.
    return ""


def _extract_reporting_period(title: str, url: str = "") -> str:
    combined = f"{title or ''} {url or ''}"
    m = re.search(r"(FY\s?\d{2,4}|1H\s?FY\s?\d{2,4}|H1\s?FY\s?\d{2,4}|20\d{2})", combined, flags=re.I)
    return m.group(1).replace(" ", "") if m else ""


def _candidate_score(candidate: Candidate) -> int:
    text = f"{candidate.title} {candidate.url}".lower()
    score = 0
    if "asx.com.au" in text or "announcements.asx.com.au" in text:
        score += 100
    if _looks_like_pdf_or_announcement(candidate.url):
        score += 60
    keyword_scores = {
        "appendix 4e": 45,
        "appendix 4d": 45,
        "annual report": 42,
        "preliminary final report": 38,
        "half year": 35,
        "half-year": 35,
        "financial report": 32,
        "report and accounts": 30,
    }
    for k, v in keyword_scores.items():
        if k in text:
            score += v
    years = [int(y) for y in re.findall(r"20\d{2}", text)]
    if years:
        score += max(years) - 2000
    for bad in LOW_VALUE_KEYWORDS:
        if bad in text:
            score -= 20
    return score


def _add_candidate(candidates: List[Candidate], title: str, url: str, source_type: str, source_page: str = "") -> None:
    url = _decode_search_href(url)
    if not url or not _is_official(url):
        return
    if not _looks_report(title, url):
        return
    c = Candidate(title=title.strip() or url, url=url, source_type=source_type, source_page=source_page)
    c.score = _candidate_score(c)
    candidates.append(c)


def _get_html(url: str, timeout: int) -> Optional[str]:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        if r.status_code >= 400:
            return None
        return r.text
    except requests.RequestException:
        return None


def _parse_links_from_page(page_url: str, html: str) -> Iterable[Tuple[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a"):
        href = a.get("href") or ""
        if not href:
            continue
        text = a.get_text(" ", strip=True) or href
        yield text, urljoin(page_url, href)


def _search_asx_historical(code: str, timeout: int) -> List[Candidate]:
    candidates: List[Candidate] = []
    current_year = datetime.utcnow().year
    years = [current_year, current_year - 1, current_year - 2]
    for year in years:
        page = f"https://www.asx.com.au/asx/v2/statistics/announcements.do?by=asxCode&asxCode={quote(code)}&timeframe=Y&year={year}"
        html = _get_html(page, timeout)
        if not html:
            continue
        for text, href in _parse_links_from_page(page, html):
            if _looks_like_pdf_or_announcement(href):
                _add_candidate(candidates, text, href, "ASX announcement", page)
    return candidates


def _search_asx_public_page(code: str, timeout: int) -> List[Candidate]:
    candidates: List[Candidate] = []
    pages = [
        f"https://www.asx.com.au/markets/trade-our-cash-market/announcements.{code.lower()}",
        f"https://www.asx.com.au/markets/trade-our-cash-market/announcements.{code.lower()}/",
    ]
    for page in pages:
        html = _get_html(page, timeout)
        if not html:
            continue
        for text, href in _parse_links_from_page(page, html):
            _add_candidate(candidates, text, href, "ASX announcement", page)
    return candidates


def _search_duckduckgo(ticker: str, company_name: str, timeout: int) -> List[Candidate]:
    candidates: List[Candidate] = []
    code = _asx_code(ticker)
    queries = [
        f'site:announcements.asx.com.au/asxpdf {code} "annual report"',
        f'site:announcements.asx.com.au/asxpdf {code} "Appendix 4E"',
        f'site:announcements.asx.com.au/asxpdf {code} "Appendix 4D"',
        f'site:asx.com.au/asxpdf {code} "financial report"',
        f'"{company_name}" "annual report" pdf investor',
        f'"{company_name}" "half year" "financial report" pdf',
    ]
    for q in queries:
        url = f"https://duckduckgo.com/html/?q={quote(q)}"
        html = _get_html(url, timeout)
        if not html:
            continue
        soup = BeautifulSoup(html, "html.parser")
        result_links = soup.select("a.result__a") or soup.find_all("a")
        for a in result_links:
            title = a.get_text(" ", strip=True)
            href = _decode_search_href(a.get("href", ""))
            if _looks_like_pdf_or_announcement(href):
                _add_candidate(candidates, title, href, "ASX announcement" if "asx.com.au" in href.lower() else "company investor centre", url)
            elif _is_official(href):
                # Open official investor page and pick report PDFs from it.
                page_html = _get_html(href, timeout)
                if not page_html:
                    continue
                for link_text, pdf_href in _parse_links_from_page(href, page_html):
                    if _looks_like_pdf_or_announcement(pdf_href):
                        _add_candidate(candidates, link_text, pdf_href, "ASX announcement" if "asx.com.au" in pdf_href.lower() else "company investor centre", href)
    return candidates


def discover_source(ticker: str, company_name: str, timeout: int = 20) -> SourceRecord:
    code = _asx_code(ticker)
    rec = SourceRecord(ticker=ticker, company_name=company_name, source_status="source not located")

    candidates: List[Candidate] = []
    discovery_steps = [
        lambda: _search_asx_historical(code, timeout),
        lambda: _search_asx_public_page(code, timeout),
        lambda: _search_duckduckgo(ticker, company_name, timeout),
    ]

    errors: List[str] = []
    for step in discovery_steps:
        try:
            candidates.extend(step())
        except Exception as exc:  # keep processing other methods
            errors.append(str(exc))

    # Deduplicate by URL and rank.
    unique = {}
    for c in candidates:
        if c.url not in unique or c.score > unique[c.url].score:
            unique[c.url] = c
    ranked = sorted(unique.values(), key=lambda c: c.score, reverse=True)

    if ranked:
        best = ranked[0]
        rec.source_status = "source located"
        rec.report_title = best.title
        rec.report_date = _extract_report_date(best.title, best.url)
        rec.reporting_period = _extract_reporting_period(best.title, best.url)
        rec.source_type = best.source_type
        rec.source_url = best.url
        rec.source_confidence = "high" if best.score >= 170 else "medium"
        rec.notes = f"Discovered automatically. Score={best.score}. Source page={best.source_page}".strip()
        return rec

    rec.source_confidence = "low"
    rec.notes = "No official report URL found automatically. " + ("; ".join(errors) if errors else "")
    return rec
