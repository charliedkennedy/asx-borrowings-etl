"""Fast, cached indexing and conservative matching of local statutory PDFs."""
from __future__ import annotations

import csv
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

import fitz


SUFFIXES = {
    "LIMITED", "LTD", "PTY", "PROPRIETARY", "HOLDINGS", "GROUP",
    "AUSTRALIA", "AUSTRALIAN", "TRUST", "REIT", "FUND", "STAPLED",
    "COMPANY", "CORPORATION", "THE",
}
ANNUAL_TYPES = {"ANNUAL_REPORT", "ANNUAL_REPORT_AND_4E"}
EXCLUDED_DESCRIPTIONS = (
    "PRESENTATION", "TRANSCRIPT", "SUSTAINABILITY", "CORPORATE GOVERNANCE",
    "NOTICE OF MEETING", "MEDIA RELEASE", "RESULTS PRESENTATION",
)
INDEX_FIELDS = [
    "full_path", "filename", "filename_ticker", "release_date",
    "financial_year", "document_description", "report_type",
    "filename_company", "acn", "file_size_bytes", "modified_time",
    "page_count", "first_pages_text", "normalised_filename_company",
    "index_error",
]
_TICKER_PATTERN = re.compile(
    r"^(?P<ticker>[A-Z0-9]{2,5})_(?P<date>\d{4}-\d{2}-\d{2})_"
    r"FY(?P<fy>\d{2}|\d{4})_(?P<description>.+)$",
    re.IGNORECASE,
)
_LEGAL_PATTERN = re.compile(
    r"^(?P<company>.+)_FY(?P<fy>\d{4})_ACN(?P<acn>\d{9})$",
    re.IGNORECASE,
)


def ratio(left: str, right: str) -> float:
    """Return a deterministic 0-100 fallback name-similarity score."""
    return SequenceMatcher(None, left, right).ratio() * 100


def normalise_name(value: str) -> str:
    value = value.upper().replace("&", " AND ")
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    words = [word for word in value.split() if word not in SUFFIXES]
    return " ".join(words).replace("GRAIN CORP", "GRAINCORP").replace(
        "BLUE SCOPE", "BLUESCOPE"
    ).replace("HEALTH CO", "HEALTHCO").replace("LONG WALE", "LONGWALE")


def classify_report_type(description: str) -> str:
    """Classify a filename description without treating presentations as reports."""
    text = re.sub(r"[_-]+", " ", description).upper()
    text = re.sub(r"\s+", " ", text).strip()
    if any(excluded in text for excluded in EXCLUDED_DESCRIPTIONS):
        return "OTHER"
    annual = "ANNUAL REPORT" in text or "FULL YEAR" in text and "REPORT" in text
    appendix_4e = "APPENDIX 4E" in text
    if annual and appendix_4e:
        return "ANNUAL_REPORT_AND_4E"
    if annual:
        return "ANNUAL_REPORT"
    if appendix_4e:
        return "APPENDIX_4E"
    if "HALF YEAR" in text or "HALF YEARLY" in text or "INTERIM FINANCIAL REPORT" in text:
        return "HALF_YEAR_REPORT"
    if "APPENDIX 4D" in text:
        return "APPENDIX_4D"
    return "OTHER"


def parse_report_filename(filename: str) -> dict[str, str]:
    """Parse ticker-prefixed reports and the older legal-name/ACN convention."""
    stem = Path(filename).stem
    ticker_match = _TICKER_PATTERN.fullmatch(stem)
    if ticker_match:
        parts = ticker_match.groupdict()
        year = parts["fy"]
        financial_year = f"20{year}" if len(year) == 2 else year
        description = parts["description"].replace("_", " ")
        try:
            release_date = date.fromisoformat(parts["date"]).isoformat()
        except ValueError:
            release_date = ""
        return {
            "filename_ticker": parts["ticker"].upper(),
            "release_date": release_date,
            "financial_year": financial_year,
            "document_description": description,
            "report_type": classify_report_type(description),
            "filename_company": "",
            "acn": "",
        }
    legal_match = _LEGAL_PATTERN.fullmatch(stem)
    if legal_match:
        parts = legal_match.groupdict()
        return {
            "filename_ticker": "", "release_date": "",
            "financial_year": parts["fy"], "document_description": "",
            "report_type": "OTHER", "filename_company": parts["company"],
            "acn": parts["acn"],
        }
    return {
        "filename_ticker": "", "release_date": "", "financial_year": "",
        "document_description": "", "report_type": "OTHER",
        "filename_company": stem, "acn": "",
    }


def parse_filename(filename: str) -> tuple[str, str, str]:
    """Backward-compatible parser for legal-name/ACN filenames."""
    parsed = parse_report_filename(filename)
    return parsed["filename_company"] or Path(filename).stem, parsed["financial_year"], parsed["acn"]


def _metadata(path: Path) -> dict:
    parsed = parse_report_filename(path.name)
    stat = path.stat()
    try:
        with fitz.open(path) as document:
            text = "\n".join(page.get_text("text") for page in list(document)[:3])
            pages = len(document)
        error = ""
    except Exception as exc:
        text, pages = "", 0
        error = f"PDF_UNREADABLE: {type(exc).__name__}: {exc}"
    return {
        "full_path": str(path), "filename": path.name, **parsed,
        "file_size_bytes": stat.st_size, "modified_time": int(stat.st_mtime),
        "page_count": pages, "first_pages_text": text[:12000],
        "normalised_filename_company": normalise_name(parsed["filename_company"]),
        "index_error": error,
    }


def build_index(pdf_root: Path, output: Path, workers: int = 8) -> list[dict]:
    """Index local PDFs, invalidating caches created with an older schema."""
    root = pdf_root.resolve(strict=True)
    if not root.is_dir():
        raise NotADirectoryError(f"PDF root is not a directory: {pdf_root}")
    output.parent.mkdir(parents=True, exist_ok=True)
    cached: dict[str, dict] = {}
    if output.exists():
        with output.open(encoding="utf-8", newline="") as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames and set(INDEX_FIELDS).issubset(reader.fieldnames):
                cached = {row["full_path"]: row for row in reader}
    paths = sorted(
        path for path in root.rglob("*")
        if path.is_file()
        and path.suffix.lower() == ".pdf"
        and path.resolve().is_relative_to(root)
    )

    def one(path: Path) -> dict:
        old = cached.get(str(path))
        stat = path.stat()
        if old and int(old["file_size_bytes"]) == stat.st_size and int(old["modified_time"]) == int(stat.st_mtime):
            return old
        return _metadata(path)

    with ThreadPoolExecutor(max_workers=min(8, max(1, workers))) as pool:
        rows = list(pool.map(one, paths))
    with output.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, INDEX_FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    return rows


def read_targets(path: Path) -> list[dict]:
    grouped: dict[str, dict] = {}
    with path.open(encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            ticker = row["ticker"].upper()
            grouped.setdefault(ticker, {
                "ticker": ticker, "target_name": row["target_name"],
                "aliases": [], "warning": "",
            })
            grouped[ticker]["aliases"] += [
                alias.strip() for alias in row.get("aliases", "").split("|") if alias.strip()
            ] + [row["target_name"]]
    return list(grouped.values())


def _identity_hit(row: dict, aliases: list[str]) -> bool:
    content = normalise_name(row.get("first_pages_text", ""))
    return any(normalise_name(alias) in content for alias in aliases)


def _exact_sort_key(row: dict, identity: bool) -> tuple:
    return (
        row.get("report_type") in ANNUAL_TYPES,
        identity,
        int(row.get("financial_year") or 0),
        row.get("release_date") or "",
        int(row.get("page_count") or 0),
    )


def match_targets(targets: Iterable[dict], index: list[dict]) -> tuple[list[dict], list[dict]]:
    """Match exact filename tickers first; use name similarity only as fallback."""
    candidates: list[dict] = []
    register: list[dict] = []
    for target in targets:
        aliases = list(dict.fromkeys(target["aliases"]))
        ticker = target["ticker"].upper()
        readable = [row for row in index if not row.get("index_error")]
        exact = [row for row in readable if row.get("filename_ticker", "").upper() == ticker]
        scored: list[tuple[float, float, bool, bool, dict, str]] = []
        if exact:
            for row in exact:
                identity = _identity_hit(row, aliases)
                annual = row.get("report_type") in ANNUAL_TYPES
                excluded = row.get("report_type") == "OTHER"
                score = min(100.0, 70 + (20 if annual else 0) + (10 if identity else 0))
                reason = "exact filename ticker" + ("; identity confirmed" if identity else "; identity not confirmed")
                if excluded:
                    reason += "; non-annual or excluded document"
                scored.append((score, 100.0, identity, True, row, reason))
            scored.sort(key=lambda item: _exact_sort_key(item[4], item[2]), reverse=True)
        else:
            norms = [normalise_name(alias) for alias in aliases]
            for row in readable:
                filename_name = row.get("normalised_filename_company", "")
                if not filename_name:
                    continue
                filename_score = max(ratio(norm, filename_name) for norm in norms)
                if filename_score < 45:
                    continue
                identity = _identity_hit(row, aliases)
                annual_text = any(marker in row.get("first_pages_text", "").upper() for marker in (
                    "ANNUAL REPORT", "FINANCIAL REPORT", "CONSOLIDATED FINANCIAL",
                ))
                score = min(100.0, filename_score * 0.65 + (25 if identity else 0) + (10 if annual_text else 0))
                scored.append((score, filename_score, identity, False, row, "fuzzy legal-name fallback"))
            scored.sort(key=lambda item: item[0], reverse=True)

        top = scored[:5]
        for rank, (overall, filename_score, identity, is_exact, row, reason) in enumerate(top, 1):
            candidates.append({
                "ticker": ticker, "target_name": target["target_name"],
                "candidate_rank": rank, "candidate_path": row["full_path"],
                "candidate_filename": row["filename"],
                "filename_ticker": row.get("filename_ticker", ""),
                "release_date": row.get("release_date", ""),
                "financial_year": row.get("financial_year", ""),
                "document_description": row.get("document_description", ""),
                "report_type": row.get("report_type", "OTHER"),
                "filename_match_score": round(filename_score, 1),
                "first_page_match_score": 100 if identity else 0,
                "overall_match_score": round(min(100.0, overall), 1),
                "matched_alias": next((alias for alias in aliases if normalise_name(alias) in normalise_name(row.get("first_pages_text", ""))), "") if identity else "",
                "acn": row.get("acn", ""), "page_count": row.get("page_count", ""),
                "reason": reason,
            })

        annual_exact = [item for item in top if item[3] and item[4].get("report_type") in ANNUAL_TYPES]
        if exact and not annual_exact:
            status, selected, confidence, reason = "AMBIGUOUS_MATCH", {}, min(100.0, top[0][0]), "only non-annual or excluded ticker files found"
        elif annual_exact:
            best = annual_exact[0]
            equally_plausible = len(annual_exact) > 1 and _exact_sort_key(best[4], best[2]) == _exact_sort_key(annual_exact[1][4], annual_exact[1][2])
            if equally_plausible:
                status, selected, confidence, reason = "AMBIGUOUS_MATCH", {}, best[0], "multiple equally plausible annual reports"
            elif best[2]:
                status, selected, confidence, reason = "MATCHED_HIGH", best[4], best[0], "exact ticker, annual report, and identity confirmed"
            else:
                status, selected, confidence, reason = "MATCHED_MEDIUM", best[4], best[0], "exact ticker and annual report; identity not fully confirmed"
        elif not top:
            status, selected, confidence, reason = "NO_LOCAL_DOCUMENT", {}, 0.0, "no credible local filename candidate"
        elif top[0][2] and (len(top) == 1 or top[0][0] - top[1][0] >= 8):
            status, selected, confidence, reason = "MATCHED_MEDIUM", top[0][4], top[0][0], "fuzzy fallback with PDF identity confirmation"
        else:
            status, selected, confidence, reason = "AMBIGUOUS_MATCH", {}, top[0][0], "fuzzy fallback is not sufficiently conclusive"

        register.append({
            "ticker": ticker, "target_name": target["target_name"],
            "aliases": " | ".join(aliases), "match_status": status,
            "match_confidence": round(min(100.0, confidence), 1),
            "selected_pdf": selected.get("full_path", ""),
            "selected_filename": selected.get("filename", ""),
            "filename_ticker": selected.get("filename_ticker", ""),
            "release_date": selected.get("release_date", ""),
            "financial_year": selected.get("financial_year", ""),
            "document_description": selected.get("document_description", ""),
            "report_type": selected.get("report_type", ""),
            "matched_legal_entity": selected.get("filename_company", ""),
            "acn": selected.get("acn", ""), "page_count": selected.get("page_count", ""),
            "match_reason": reason,
            "alternative_candidates": " | ".join(item[4]["filename"] for item in top if item[4] is not selected),
            "validation_warning": target.get("warning", ""),
        })
    return candidates, register
