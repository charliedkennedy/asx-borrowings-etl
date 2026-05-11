import re
from datetime import datetime
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

from .schemas import SourceRecord

REPORT_KEYWORDS = [
    "annual report", "half year", "half-year", "appendix 4d", "appendix 4e",
    "preliminary final", "financial report", "results"
]


def _is_official(url: str) -> bool:
    u = (url or "").lower()
    return "asx.com.au" in u or ".com.au" in u


def _extract_date(text: str) -> str:
    m = re.search(r"(20\d{2})", text or "")
    return m.group(1) if m else ""


def discover_source(ticker: str, company_name: str, timeout: int = 20) -> SourceRecord:
    rec = SourceRecord(ticker=ticker, company_name=company_name, source_status="source not located")
    queries = [
        f"site:asx.com.au {ticker} announcement annual report pdf",
        f"site:asx.com.au {ticker} appendix 4e pdf",
        f"{company_name} investor centre annual report pdf",
    ]
    headers = {"User-Agent": "Mozilla/5.0"}

    candidates = []
    for q in queries:
        try:
            resp = requests.get(f"https://duckduckgo.com/html/?q={quote(q)}", headers=headers, timeout=timeout)
            soup = BeautifulSoup(resp.text, "html.parser")
            for a in soup.select("a.result__a"):
                title = a.get_text(" ", strip=True)
                href = a.get("href", "")
                t = title.lower()
                if not any(k in t for k in REPORT_KEYWORDS):
                    continue
                if not _is_official(href):
                    continue
                candidates.append((title, href))
        except Exception as e:
            rec.notes = f"discovery error: {e}"

    if candidates:
        title, href = candidates[0]
        rec.source_status = "source located"
        rec.report_title = title
        rec.report_date = datetime.utcnow().date().isoformat()
        rec.reporting_period = _extract_date(title)
        rec.source_type = "ASX announcement" if "asx.com.au" in href.lower() else "company investor centre"
        rec.source_url = href
        rec.source_confidence = "medium"
        return rec

    rec.report_date = datetime.utcnow().date().isoformat()
    rec.notes = rec.notes or "No official report URL found automatically; manual source URL required."
    return rec
