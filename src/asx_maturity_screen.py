"""Local-only ASX debt maturity screen. It never retrieves source PDFs online."""
from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Callable, Literal

import fitz
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, ConfigDict, Field

from .local_pdf_index import build_index, match_targets, read_targets


PILOT = ("IDX", "MTS", "EVT", "XRO", "BSL", "CHC", "DMP", "GNC", "S32", "TAH")
ACCEPTED = {"MATCHED_HIGH", "MATCHED_MEDIUM"}
SUCCESS_STATUSES = {"EXTRACTED", "EXTRACTED_WITH_FLAGS"}
HALVES = ("2H26", "1H27", "2H27", "1H28", "2H28", "FY29+")
KEYWORDS = (
    "borrowings", "loans and borrowings", "interest-bearing liabilities",
    "interest bearing", "financing facilities", "debt facilities",
    "bank facilities", "liquidity", "capital management", "maturity",
    "maturities", "contractual maturities", "current borrowings",
    "non-current borrowings", "cash and cash equivalents", "net debt",
    "lease liabilities",
)


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class MoneyValue(StrictModel):
    value_m: float | None
    currency: str | None
    source_page: int | None
    source_quote_or_evidence: str | None


class Facility(StrictModel):
    ticker: str
    facility_or_instrument_name: str
    lender_or_market: str | None
    instrument_type: str | None
    currency: str | None
    facility_limit: MoneyValue | None
    drawn_amount: MoneyValue | None
    undrawn_amount: MoneyValue | None
    maturity_date: str | None
    maturity_description: str | None
    secured_or_unsecured: str | None
    current_or_non_current: str | None
    source_page: int | None
    source_quote_or_evidence: str | None
    confidence: float = Field(ge=0, le=1)


class DisclosureBucket(StrictModel):
    ticker: str
    bucket_label: str
    period_start: str | None
    period_end: str | None
    amount: MoneyValue | None
    currency: str | None
    source_page: int | None
    source_quote_or_evidence: str | None
    confidence: float = Field(ge=0, le=1)


class ExtractionPayload(StrictModel):
    ticker: str
    company_name: str
    matched_legal_entity: str | None
    financial_year: str | None
    balance_date: str | None
    reporting_currency: str | None
    reporting_unit: str | None
    gross_debt_excluding_leases: MoneyValue | None
    current_borrowings_excluding_leases: MoneyValue | None
    non_current_borrowings_excluding_leases: MoneyValue | None
    undrawn_committed_headroom: MoneyValue | None
    cash_and_cash_equivalents: MoneyValue | None
    net_debt_excluding_leases: MoneyValue | None
    lease_liabilities: MoneyValue | None
    facilities: list[Facility]
    disclosure_buckets: list[DisclosureBucket]
    extraction_confidence: float = Field(ge=0, le=1)
    extraction_status: str
    model_used: str
    source_pages: list[int]
    extraction_notes: str
    validation_flags: list[str]
    profile_quality: Literal["FACILITY_DATED", "BUCKET_INFERRED", "SPLIT_ONLY"] | None
    leases_apparently_included: bool


EXTRACTION_PROMPT = """Extract debt and maturity information from selected pages of an Australian statutory report.
Return only facts supported by the supplied page text. Never guess. All monetary values must be converted to millions using the disclosed reporting unit, retain their currency, source page, and short source evidence. Use null for undisclosed fields. Exclude lease liabilities, derivatives not explicitly reported as borrowings, and trade payables from debt and maturities. Keep lease liabilities separately. Distinguish committed undrawn headroom from facility limits. Determine the balance date and financial year primarily from report contents. Preserve exact facilities separately from broad maturity buckets. Do not double count them. Dates must use YYYY-MM-DD when reliably disclosed. Page labels in the supplied text are the source page numbers."""


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--pdf-root", required=True)
    result.add_argument("--targets", required=True)
    result.add_argument("--output", default="outputs/asx_maturity_screen.xlsx")
    result.add_argument("--pilot", action="store_true")
    result.add_argument("--match-only", action="store_true")
    result.add_argument("--ticker", type=str.upper)
    result.add_argument("--force", action="store_true")
    result.add_argument("--model", default="gpt-4.1-mini")
    return result


def to_millions(value: float | None, reporting_unit: str | None) -> float | None:
    if value is None:
        return None
    unit = (reporting_unit or "").upper().replace(" ", "")
    if unit in {"$'000", "$000", "000", "THOUSAND", "THOUSANDS"}:
        return value / 1000
    if unit in {"$", "UNIT", "UNITS"}:
        return value / 1_000_000
    return value


def calendar_half(value: str | None) -> str | None:
    parsed = parse_date(value)
    if not parsed:
        return None
    if parsed.year >= 2029:
        return "FY29+"
    key = f"{'1H' if parsed.month <= 6 else '2H'}{str(parsed.year)[-2:]}"
    return key if key in HALVES else None


def allocate_exact(amount: float | None, maturity_date: str | None) -> dict[str, float | None]:
    result = {half: None for half in HALVES}
    half = calendar_half(maturity_date)
    if amount is not None and half:
        result[half] = amount
    return result


def allocate_bucket(
    amount: float | None, start_month: int | None, end_month: int | None,
    balance_year: int = 2026, balance_month: int = 7,
) -> dict[str, float | None]:
    result = {half: None for half in HALVES}
    if amount is None or start_month is None or end_month is None or end_month <= start_month:
        return result
    for month_offset in range(start_month, end_month):
        year = balance_year + (balance_month - 1 + month_offset) // 12
        month = (balance_month - 1 + month_offset) % 12 + 1
        half = calendar_half(f"{year}-{month:02d}-01")
        if half:
            result[half] = (result[half] or 0) + amount / (end_month - start_month)
    return result


def allocate_bucket_dates(
    amount: float | None, period_start: str | None, period_end: str | None,
) -> dict[str, float | None]:
    result = {half: None for half in HALVES}
    start, end = parse_date(period_start), parse_date(period_end)
    if amount is None or not start or not end or end < start:
        return result
    months: list[date] = []
    cursor = date(start.year, start.month, 1)
    finish = date(end.year, end.month, 1)
    while cursor <= finish:
        months.append(cursor)
        cursor = date(cursor.year + (cursor.month == 12), 1 if cursor.month == 12 else cursor.month + 1, 1)
    if not months:
        return result
    for month in months:
        half = calendar_half(month.isoformat())
        if half:
            result[half] = (result[half] or 0) + amount / len(months)
    return result


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value[:10])
    except (TypeError, ValueError):
        return None


def money(value: MoneyValue | None) -> float | None:
    return value.value_m if value else None


def select_relevant_pages(pdf_path: Path, maximum: int = 16) -> tuple[list[int], list[str]]:
    """Extract all text locally, then retain keyword pages and their neighbours."""
    with fitz.open(pdf_path) as document:
        texts = [page.get_text("text") for page in document]
    scores: list[tuple[int, int]] = []
    for index, text in enumerate(texts):
        lowered = text.lower()
        score = sum(lowered.count(keyword) for keyword in KEYWORDS)
        if score:
            score += min(5, sum(character.isdigit() for character in text) // 30)
            scores.append((score, index))
    if not scores:
        return [], []
    selected: set[int] = set()
    for _, index in sorted(scores, reverse=True):
        for neighbour in (index - 1, index, index + 1):
            if 0 <= neighbour < len(texts):
                selected.add(neighbour)
        if len(selected) >= maximum:
            break
    ranked = sorted(selected, key=lambda item: (-next((score for score, index in scores if index == item), 0), item))[:maximum]
    pages = sorted(ranked)
    return [page + 1 for page in pages], [texts[page] for page in pages]


def extract_with_openai(
    ticker: str, company_name: str, matched_legal_entity: str | None,
    filename_financial_year: str | None, source_pages: list[int], page_texts: list[str],
    model: str, api_key: str,
) -> tuple[ExtractionPayload, dict]:
    client = OpenAI(api_key=api_key)
    page_content = "\n\n".join(
        f"--- PDF SOURCE PAGE {page} ---\n{text}" for page, text in zip(source_pages, page_texts)
    )
    completion = client.beta.chat.completions.parse(
        model=model,
        messages=[
            {"role": "system", "content": EXTRACTION_PROMPT},
            {"role": "user", "content": (
                f"Ticker: {ticker}\nCompany: {company_name}\nMatched legal entity: {matched_legal_entity or ''}\n"
                f"Filename financial year (supporting only): {filename_financial_year or ''}\n\n{page_content}"
            )},
        ],
        response_format=ExtractionPayload,
    )
    parsed = completion.choices[0].message.parsed
    if parsed is None:
        raise ValueError("Model returned no parsed extraction")
    usage = completion.usage
    return parsed, {
        "input_tokens": getattr(usage, "prompt_tokens", None),
        "output_tokens": getattr(usage, "completion_tokens", None),
    }


def validate_extraction(payload: ExtractionPayload, filename_year: str | None) -> list[str]:
    flags = list(dict.fromkeys(payload.validation_flags))
    gross = money(payload.gross_debt_excluding_leases)
    current = money(payload.current_borrowings_excluding_leases)
    non_current = money(payload.non_current_borrowings_excluding_leases)
    if gross is not None and current is not None and non_current is not None:
        if abs(current + non_current - gross) / max(abs(gross), 1e-9) > 0.02:
            flags.append("CURRENT_NON_CURRENT_RECON_FAIL")
    drawn = [money(facility.drawn_amount) for facility in payload.facilities]
    drawn_total = sum(value for value in drawn if value is not None)
    if gross is not None and drawn and any(value is not None for value in drawn):
        if abs(drawn_total - gross) / max(abs(gross), 1e-9) > 0.07:
            flags.append("FACILITY_RECON_FAIL")
    money_objects = [
        payload.gross_debt_excluding_leases, payload.current_borrowings_excluding_leases,
        payload.non_current_borrowings_excluding_leases, payload.undrawn_committed_headroom,
        payload.cash_and_cash_equivalents, payload.net_debt_excluding_leases,
        payload.lease_liabilities,
    ]
    monetary_values = [money(value) for value in money_objects]
    for facility in payload.facilities:
        money_objects.extend((facility.facility_limit, facility.drawn_amount, facility.undrawn_amount))
        monetary_values.extend((money(facility.facility_limit), money(facility.drawn_amount), money(facility.undrawn_amount)))
        limit, undrawn = money(facility.facility_limit), money(facility.undrawn_amount)
        if limit is not None and undrawn is not None and undrawn > limit:
            flags.append("UNDRAWN_EXCEEDS_LIMIT")
    for bucket in payload.disclosure_buckets:
        money_objects.append(bucket.amount)
        monetary_values.append(money(bucket.amount))
    if any(value is not None and value < 0 for value in monetary_values):
        flags.append("NEGATIVE_DEBT_OR_FACILITY_AMOUNT")
    if gross is not None and not 0 <= gross <= 25_000:
        flags.append("GROSS_DEBT_OUT_OF_RANGE")
    balance_date = parse_date(payload.balance_date)
    if balance_date:
        for facility in payload.facilities:
            maturity = parse_date(facility.maturity_date)
            if maturity and maturity < balance_date:
                flags.append("MATURITY_BEFORE_BALANCE_DATE")
    if not payload.reporting_currency:
        flags.append("MISSING_REPORTING_CURRENCY")
    if not payload.source_pages or any(
        value is not None and value.value_m is not None and value.source_page is None
        for value in money_objects
    ):
        flags.append("MISSING_SOURCE_PAGES")
    if payload.leases_apparently_included:
        flags.append("LEASES_APPARENTLY_INCLUDED")
    if filename_year and balance_date and filename_year != str(balance_date.year):
        flags.append("FILENAME_YEAR_BALANCE_DATE_CONFLICT")
    if payload.extraction_confidence < 0.60:
        flags.append("LOW_CONFIDENCE")
    if payload.facilities and any(
        money(facility.drawn_amount) is not None and not facility.maturity_date
        for facility in payload.facilities
    ):
        flags.append("MATURITY_PROFILE_INCOMPLETE")
    return list(dict.fromkeys(flags))


def maturity_grid(payload: ExtractionPayload) -> tuple[dict[str, float | None], str]:
    exact = {half: None for half in HALVES}
    inferred = {half: None for half in HALVES}
    exact_found = False
    for facility in payload.facilities:
        amount = money(facility.drawn_amount)
        allocation = allocate_exact(amount, facility.maturity_date)
        for half, value in allocation.items():
            if value is not None:
                exact[half] = (exact[half] or 0) + value
                exact_found = True
    if exact_found:
        quality = "FACILITY_DATED"
    else:
        for bucket in payload.disclosure_buckets:
            allocation = allocate_bucket_dates(money(bucket.amount), bucket.period_start, bucket.period_end)
            for half, value in allocation.items():
                if value is not None:
                    inferred[half] = (inferred[half] or 0) + value
        if any(value is not None for value in inferred.values()):
            quality = "BUCKET_INFERRED"
        elif money(payload.current_borrowings_excluding_leases) is not None or money(payload.non_current_borrowings_excluding_leases) is not None:
            quality = "SPLIT_ONLY"
        else:
            quality = payload.profile_quality or "SPLIT_ONLY"
    grid: dict[str, float | None] = {}
    for half in HALVES:
        grid[f"{half} exact"] = exact[half]
        grid[f"{half} inferred"] = inferred[half]
        grid[f"{half} total"] = sum(value for value in (exact[half], inferred[half]) if value is not None) if exact[half] is not None or inferred[half] is not None else None
    return grid, quality


def write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]) if rows else ["ticker"])
        writer.writeheader()
        writer.writerows(rows)


def load_results(path: Path) -> dict[str, dict]:
    latest: dict[str, dict] = {}
    if not path.exists():
        return latest
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            record = json.loads(line)
        except json.JSONDecodeError:
            continue
        latest[record["ticker"]] = record
    return latest


def append_result(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def document_fingerprint(selected_pdf: str) -> dict:
    stat = Path(selected_pdf).stat()
    return {"selected_pdf": selected_pdf, "file_size": stat.st_size, "modified_time": int(stat.st_mtime)}


def build_record(
    register: dict, payload: ExtractionPayload | None, status: str, flags: list[str],
    grid: dict[str, float | None], profile_quality: str | None, run_log: dict,
    error: str = "",
) -> dict:
    return {
        "ticker": register["ticker"], "target_name": register["target_name"],
        "document_match": register, "document_fingerprint": document_fingerprint(register["selected_pdf"]) if register.get("selected_pdf") else None,
        "selected_pages": payload.source_pages if payload else [],
        "extraction": payload.model_dump(mode="json") if payload else None,
        "status": status, "validation_flags": flags, "maturity_grid": grid,
        "profile_quality": profile_quality, "run_log": run_log, "error": error,
    }


def summary_row(record: dict) -> dict:
    match = record["document_match"]
    extraction = record.get("extraction") or {}
    def amount(field: str) -> float | None:
        value = extraction.get(field)
        return value.get("value_m") if value else None
    row = {
        "Ticker": record["ticker"], "Company name": record["target_name"],
        "Matched legal entity": extraction.get("matched_legal_entity") or match.get("matched_legal_entity", ""),
        "Match status": match["match_status"], "Match confidence": match["match_confidence"],
        "Extraction status": record["status"], "Extraction confidence": extraction.get("extraction_confidence"),
        "Model used": extraction.get("model_used"), "Financial year": extraction.get("financial_year") or match.get("financial_year"),
        "Report type": match.get("report_type"), "Balance date": extraction.get("balance_date"),
        "Reporting currency": extraction.get("reporting_currency"),
        "Gross debt ex leases": amount("gross_debt_excluding_leases"),
        "Current borrowings": amount("current_borrowings_excluding_leases"),
        "Non-current borrowings": amount("non_current_borrowings_excluding_leases"),
        "Undrawn committed headroom": amount("undrawn_committed_headroom"),
        "Cash": amount("cash_and_cash_equivalents"), "Net debt": amount("net_debt_excluding_leases"),
        **record.get("maturity_grid", {}),
        "2H27 maturity total": record.get("maturity_grid", {}).get("2H27 total"),
        "2H27 maturity flag": "Y" if record.get("maturity_grid", {}).get("2H27 total") not in (None, 0) else "",
        "Profile quality": record.get("profile_quality"), "Selected PDF": match.get("selected_pdf"),
        "Source pages": ", ".join(map(str, record.get("selected_pages", []))),
        "Validation flags": " | ".join(record.get("validation_flags", [])),
        "Extraction notes": extraction.get("extraction_notes") or record.get("error", ""),
    }
    return row


SUMMARY_HEADERS = [
    "Ticker", "Company name", "Matched legal entity", "Match status", "Match confidence",
    "Extraction status", "Extraction confidence", "Model used", "Financial year", "Report type",
    "Balance date", "Reporting currency", "Gross debt ex leases", "Current borrowings",
    "Non-current borrowings", "Undrawn committed headroom", "Cash", "Net debt",
] + [f"{half} {kind}" for half in HALVES for kind in ("exact", "inferred", "total")] + [
    "2H27 maturity total", "2H27 maturity flag", "Profile quality", "Selected PDF",
    "Source pages", "Validation flags", "Extraction notes",
]
FACILITY_HEADERS = [
    "company_name", "selected_pdf", "ticker", "facility_or_instrument_name",
    "lender_or_market", "instrument_type", "currency", "maturity_date",
    "maturity_description", "secured_or_unsecured", "current_or_non_current",
    "source_page", "source_quote_or_evidence", "confidence", "facility_limit_m",
    "facility_limit_source_page", "drawn_amount_m", "drawn_amount_source_page",
    "undrawn_amount_m", "undrawn_amount_source_page",
]
BUCKET_HEADERS = [
    "company_name", "selected_pdf", "ticker", "bucket_label", "period_start",
    "period_end", "currency", "source_page", "source_quote_or_evidence",
    "confidence", "amount_m", "amount_source_page",
]
EXCEPTION_HEADERS = [
    "ticker", "company", "status", "validation_flags", "error", "selected_pdf",
]
RUN_LOG_HEADERS = [
    "ticker", "selected_pdf", "start_time", "finish_time", "elapsed_seconds",
    "extraction_status", "retry_count", "model_used", "source_pages",
    "error_message", "input_tokens", "output_tokens",
]


def detail_rows(records: list[dict]) -> tuple[list[dict], list[dict]]:
    facilities: list[dict] = []
    buckets: list[dict] = []
    for record in records:
        extraction = record.get("extraction") or {}
        for item in extraction.get("facilities", []):
            row = {"company_name": record["target_name"], "selected_pdf": record["document_match"].get("selected_pdf"), **item}
            for field in ("facility_limit", "drawn_amount", "undrawn_amount"):
                value = row.pop(field, None)
                row[f"{field}_m"] = value.get("value_m") if value else None
                row[f"{field}_source_page"] = value.get("source_page") if value else None
            facilities.append(row)
        for item in extraction.get("disclosure_buckets", []):
            row = {"company_name": record["target_name"], "selected_pdf": record["document_match"].get("selected_pdf"), **item}
            value = row.pop("amount", None)
            row["amount_m"] = value.get("value_m") if value else None
            row["amount_source_page"] = value.get("source_page") if value else None
            buckets.append(row)
    return facilities, buckets


def workbook(path: Path, records: list[dict], register: list[dict]) -> None:
    summary = [summary_row(record) for record in records]
    facilities, buckets = detail_rows(records)
    exceptions: list[dict] = []
    for record in records:
        if record["status"] not in SUCCESS_STATUSES or record.get("validation_flags"):
            exceptions.append({
                "ticker": record["ticker"], "company": record["target_name"],
                "status": record["status"], "validation_flags": " | ".join(record.get("validation_flags", [])),
                "error": record.get("error", ""), "selected_pdf": record["document_match"].get("selected_pdf", ""),
            })
    run_logs = [record.get("run_log", {}) for record in records]
    sheets = [
        ("Summary", summary, SUMMARY_HEADERS),
        ("Facilities & Instruments", facilities, FACILITY_HEADERS),
        ("Disclosure Buckets", buckets, BUCKET_HEADERS),
        ("Document Register", register, list(register[0]) if register else ["ticker"]),
        ("Exceptions", exceptions, EXCEPTION_HEADERS),
        ("Run Log", run_logs, RUN_LOG_HEADERS),
    ]
    book = Workbook()
    book.remove(book.active)
    for name, rows, headers in sheets:
        sheet = book.create_sheet(name)
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(header) for header in headers])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in range(1, len(headers) + 1):
            sheet.column_dimensions[get_column_letter(column)].width = min(45, max(12, len(headers[column - 1]) + 2))
        if name == "Summary":
            inferred_fill = PatternFill("solid", fgColor="D9D9D9")
            for column, header in enumerate(headers, 1):
                if header.endswith(" inferred"):
                    for cell in sheet.iter_cols(min_col=column, max_col=column, min_row=2):
                        cell[0].fill = inferred_fill
                if header == "Selected PDF":
                    for row_number, record in enumerate(rows, 2):
                        selected = record.get(header)
                        if selected:
                            sheet.cell(row_number, column).hyperlink = Path(selected).as_uri()
                            sheet.cell(row_number, column).style = "Hyperlink"
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)


def run_local(args: argparse.Namespace, extractor: Callable = extract_with_openai) -> int:
    output_path = Path(args.output)
    output_dir = output_path.parent
    pdf_root = Path(args.pdf_root).resolve(strict=True)
    indexed = build_index(pdf_root, output_dir / "pdf_index.csv")
    targets = read_targets(Path(args.targets))
    if args.pilot:
        targets = [target for target in targets if target["ticker"] in PILOT]
    if args.ticker:
        targets = [target for target in targets if target["ticker"] == args.ticker]
    candidates, register = match_targets(targets, indexed)
    write_csv(output_dir / "document_match_candidates.csv", candidates)
    write_csv(output_dir / "document_register.csv", register)
    if args.match_only:
        print(f"Local matching complete: {output_dir / 'document_register.csv'}")
        return 0
    accepted = [item for item in register if item["match_status"] in ACCEPTED]
    api_key = os.getenv("OPENAI_API_KEY")
    if accepted and not api_key:
        print("OPENAI_API_KEY is required after matching when accepted PDFs need extraction.", file=sys.stderr)
        return 2
    results_path = output_dir / "results.jsonl"
    latest = load_results(results_path)
    for item in register:
        if item["match_status"] not in ACCEPTED:
            record = build_record(item, None, item["match_status"], [], {}, None, {
                "ticker": item["ticker"], "selected_pdf": "", "extraction_status": item["match_status"],
            }, item["match_reason"])
            append_result(results_path, record)
            latest[item["ticker"]] = record
            continue
        selected = Path(item["selected_pdf"]).resolve(strict=True)
        if not selected.is_relative_to(pdf_root):
            raise ValueError(f"Selected PDF is outside --pdf-root: {selected}")
        fingerprint = document_fingerprint(str(selected))
        previous = latest.get(item["ticker"])
        if not args.force and previous and previous.get("status") in SUCCESS_STATUSES and previous.get("document_fingerprint") == fingerprint:
            continue
        started = datetime.now(timezone.utc)
        monotonic_start = time.monotonic()
        retries = 0
        payload = None
        usage: dict = {}
        error = ""
        source_pages: list[int] = []
        try:
            source_pages, page_texts = select_relevant_pages(selected)
            if not source_pages:
                raise ValueError("No debt-related source pages were identified")
            for attempt in range(3):
                retries = attempt
                try:
                    candidate, candidate_usage = extractor(
                        item["ticker"], item["target_name"], item.get("matched_legal_entity"),
                        item.get("financial_year"), source_pages, page_texts, args.model, api_key,
                    )
                    payload, usage = candidate, candidate_usage
                    if payload.extraction_confidence < 0.60 and attempt < 2:
                        error = "Extraction confidence below 0.60; retrying"
                        continue
                    break
                except Exception as exc:
                    error = f"{type(exc).__name__}: {exc}"
                    if attempt == 2:
                        raise
            if payload:
                payload.source_pages = sorted(set(payload.source_pages) | set(source_pages))
                payload.model_used = args.model
                flags = validate_extraction(payload, item.get("financial_year"))
                grid, quality = maturity_grid(payload)
                status = "EXTRACTED_WITH_FLAGS" if flags else "EXTRACTED"
            else:
                raise ValueError("Extraction returned no payload")
        except Exception as exc:
            error = error or f"{type(exc).__name__}: {exc}"
            flags, grid, quality, status = [], {}, None, "EXTRACTION_FAILED"
        finished = datetime.now(timezone.utc)
        run_log = {
            "ticker": item["ticker"], "selected_pdf": str(selected),
            "start_time": started.isoformat(), "finish_time": finished.isoformat(),
            "elapsed_seconds": round(time.monotonic() - monotonic_start, 3),
            "extraction_status": status, "retry_count": retries,
            "model_used": args.model, "source_pages": ", ".join(map(str, source_pages)),
            "error_message": error, **usage,
        }
        record = build_record(item, payload, status, flags, grid, quality, run_log, error)
        append_result(results_path, record)
        latest[item["ticker"]] = record
    records = [latest[item["ticker"]] for item in register if item["ticker"] in latest]
    workbook(output_path, records, register)
    print(f"Workbook: {output_path}")
    return 0


def main_for_args(argv=None, extractor: Callable = extract_with_openai) -> int:
    return run_local(parser().parse_args(argv), extractor=extractor)


def main() -> int:
    return main_for_args()


if __name__ == "__main__":
    raise SystemExit(main())
