import re
from typing import Dict, Iterable, Optional

from .maturity_mapper import blank_maturity


MONTHS = {
    "january": "01", "february": "02", "march": "03", "april": "04", "may": "05", "june": "06",
    "july": "07", "august": "08", "september": "09", "october": "10", "november": "11", "december": "12",
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "jun": "06", "jul": "07", "aug": "08", "sep": "09", "sept": "09", "oct": "10", "nov": "11", "dec": "12",
}


def normalise_to_millions(raw: str, unit_hint: str = ""):
    if raw is None:
        return None
    s = str(raw).replace(",", "").replace("$", "").strip()
    if s in {"", "-", "–"}:
        return 0.0
    try:
        v = float(s)
    except Exception:
        return None
    u = (unit_hint or "").lower()
    if "billion" in u or "bn" in u:
        return round(v * 1000.0, 3)
    if "000" in u or "thousand" in u:
        return round(v / 1000.0, 3)
    if "$m" in u or "million" in u or "millions" in u:
        return round(v, 3)
    return round(v, 3)


def _detect_unit(text: str) -> str:
    sample = text[:25000].lower()
    unit_patterns = [
        r"amounts?\s+(?:are\s+)?(?:shown|presented|stated)?\s*(?:in)?\s*(a\$|aud|\$)?\s*'?000",
        r"\$'?000",
        r"a\$'?000",
        r"in\s+thousands",
        r"amounts?\s+(?:are\s+)?(?:shown|presented|stated)?\s*(?:in)?\s*(a\$|aud|\$)?\s*m(?:illion)?s?",
        r"\$m",
        r"in\s+millions",
    ]
    for p in unit_patterns:
        m = re.search(p, sample)
        if m:
            return m.group(0)
    return ""


def _extract_currency(text: str) -> str:
    sample = text[:30000].lower()
    if any(x in sample for x in ["new zealand dollars", "nzd", "nz$"]):
        return "NZD"
    if any(x in sample for x in ["united states dollars", "usd", "us$"]):
        return "USD"
    if any(x in sample for x in ["australian dollars", "aud", "a$"]):
        return "AUD"
    if "$" in sample:
        return "AUD"
    return ""


def _extract_report_date(text: str) -> str:
    sample = text[:50000]
    patterns = [
        r"(?:as at|for the (?:half[- ]year|year) ended)\s+(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})",
        r"(?:as at|for the (?:half[- ]year|year) ended)\s+(\d{1,2})[\-/](\d{1,2})[\-/](20\d{2})",
    ]
    for p in patterns:
        m = re.search(p, sample, flags=re.I)
        if not m:
            continue
        if m.group(2).isdigit():
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            month_key = m.group(2).lower()[:4] if m.group(2).lower().startswith("sept") else m.group(2).lower()[:3]
            month = int(MONTHS.get(month_key, "0"))
            day, year = int(m.group(1)), int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{year:04d}-{month:02d}-{day:02d}"
    return ""


def _number_pattern() -> str:
    return r"(\(?-?\d[\d,]*(?:\.\d+)?\)?|[-–])"


def _clean_number(raw: str) -> str:
    return str(raw).replace("(", "-").replace(")", "").replace("–", "-").strip()


def _search_amount(text: str, labels: Iterable[str], unit_hint: str) -> Optional[float]:
    number = _number_pattern()
    # Common extracted-table pattern: label, optional words/spaces, number.
    for label in labels:
        patterns = [
            rf"{label}[^\n\d\-–]{{0,120}}{number}",
            rf"{label}[\s\S]{{0,180}}?{number}",
        ]
        for p in patterns:
            for m in re.finditer(p, text, flags=re.I):
                context = m.group(0).lower()
                if any(x in context for x in ["lease", "right-of-use", "trade", "payable"]):
                    continue
                value = normalise_to_millions(_clean_number(m.group(1)), unit_hint)
                if value is not None:
                    return value
    return None


def _has_nil_borrowings(lower: str) -> bool:
    nil_phrases = [
        "no borrowings",
        "nil borrowings",
        "no interest-bearing liabilities",
        "no loans and borrowings",
        "no bank debt",
        "debt free",
    ]
    return any(p in lower for p in nil_phrases)


def extract_borrowings_from_text(text: str) -> Dict:
    lower = text.lower()
    unit_hint = _detect_unit(text)
    out = {
        "report_date": _extract_report_date(text),
        "currency": _extract_currency(text),
        "current": "not disclosed",
        "non_current": "not disclosed",
        "total": "not disclosed",
        "maturity": blank_maturity(),
        "status": "report located but borrowings not found",
        "notes": "",
        "manual_review": "Y",
    }

    if _has_nil_borrowings(lower):
        out["current"] = out["non_current"] = out["total"] = 0.0
        out["maturity"] = blank_maturity(0.0)
        out["status"] = "nil borrowings"
        out["manual_review"] = "N"
        return out

    current_labels = [
        r"current\s+borrowings",
        r"current\s+loans\s+and\s+borrowings",
        r"borrowings\s+current",
        r"current\s+interest[-\s]?bearing\s+liabilities",
        r"current\s+bank\s+loans?",
    ]
    non_current_labels = [
        r"non[-\s]?current\s+borrowings",
        r"non[-\s]?current\s+loans\s+and\s+borrowings",
        r"borrowings\s+non[-\s]?current",
        r"non[-\s]?current\s+interest[-\s]?bearing\s+liabilities",
        r"non[-\s]?current\s+bank\s+loans?",
    ]
    total_labels = [
        r"total\s+borrowings",
        r"total\s+loans\s+and\s+borrowings",
        r"total\s+interest[-\s]?bearing\s+liabilities",
        r"borrowings\s+total",
        r"bank\s+debt\s+total",
    ]

    current = _search_amount(text, current_labels, unit_hint)
    non_current = _search_amount(text, non_current_labels, unit_hint)
    total = _search_amount(text, total_labels, unit_hint)

    if current is not None:
        out["current"] = current
    if non_current is not None:
        out["non_current"] = non_current
    if total is not None:
        out["total"] = total

    if isinstance(out["current"], (int, float)) and isinstance(out["non_current"], (int, float)) and out["total"] == "not disclosed":
        out["total"] = round(out["current"] + out["non_current"], 3)

    if isinstance(out["current"], (int, float)) and isinstance(out["non_current"], (int, float)):
        out["status"] = "current/non-current extracted, maturity not disclosed"
        out["manual_review"] = "N"
    elif isinstance(out["total"], (int, float)):
        out["status"] = "current/non-current extracted, maturity not disclosed"
        out["notes"] += "Total borrowings extracted but current/non-current split not found. "
        out["manual_review"] = "Y"

    if unit_hint:
        out["notes"] += f"Unit detected: {unit_hint}. "
    else:
        out["notes"] += "Unit not confidently detected; numbers may require manual review. "
        if any(isinstance(out[k], (int, float)) for k in ["current", "non_current", "total"]):
            out["manual_review"] = "Y"

    if any(k in lower for k in ["lease liabilities", "including leases", "right-of-use"]):
        out["notes"] += "PDF contains lease liability references; confirm leases are excluded. "
        if out["status"] != "nil borrowings":
            out["manual_review"] = "Y"

    return out
