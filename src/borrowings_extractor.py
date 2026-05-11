import re
from typing import Dict, Tuple
from .maturity_mapper import blank_maturity


def normalise_to_millions(raw: str, unit_hint: str = ""):
    s = raw.replace(",", "").strip()
    try:
        v = float(s)
    except Exception:
        return None
    u = unit_hint.lower()
    if "000" in u:
        return round(v / 1000.0, 3)
    if "$m" in u or "million" in u:
        return round(v, 3)
    if "$'000" in u or "thousand" in u:
        return round(v / 1000.0, 3)
    if "bn" in u or "billion" in u:
        return round(v * 1000.0, 3)
    return round(v, 3)


def extract_borrowings_from_text(text: str) -> Dict:
    lower = text.lower()
    out = {
        "report_date": "",
        "currency": "AUD" if "aud" in lower or "a$" in lower else "",
        "current": "not disclosed",
        "non_current": "not disclosed",
        "total": "not disclosed",
        "maturity": blank_maturity(),
        "status": "report located but borrowings not found",
        "notes": "",
        "manual_review": "Y",
    }
    if "no borrowings" in lower or "nil borrowings" in lower:
        out["current"] = out["non_current"] = out["total"] = 0.0
        out["maturity"] = blank_maturity(0.0)
        out["status"] = "nil borrowings"
        out["manual_review"] = "N"
        return out

    unit_hint = ""
    m_unit = re.search(r"\$\s?\(?('?0{3}|m)\)?", lower)
    if m_unit:
        unit_hint = m_unit.group(0)

    def grab(pattern):
        m = re.search(pattern, lower)
        if not m:
            return None
        return normalise_to_millions(m.group(1), unit_hint)

    out["current"] = grab(r"current borrowings[^\d]{0,40}([\d,]+(?:\.\d+)?)") or out["current"]
    out["non_current"] = grab(r"non[-\s]?current borrowings[^\d]{0,40}([\d,]+(?:\.\d+)?)") or out["non_current"]
    out["total"] = grab(r"total borrowings[^\d]{0,40}([\d,]+(?:\.\d+)?)") or out["total"]

    if out["current"] != "not disclosed" and out["non_current"] != "not disclosed" and out["total"] == "not disclosed":
        out["total"] = round(out["current"] + out["non_current"], 3)

    if out["current"] != "not disclosed" and out["non_current"] != "not disclosed":
        out["status"] = "current/non-current extracted, maturity not disclosed"
        out["manual_review"] = "N"

    if any(k in lower for k in ["lease liabilities", "including leases"]):
        out["notes"] += "Possible lease liabilities included. "
        out["manual_review"] = "Y"

    return out
