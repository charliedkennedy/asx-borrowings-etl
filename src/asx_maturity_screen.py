"""Local-only ASX debt maturity screen. It never retrieves source PDFs online."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Callable, Literal

import fitz
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, ConfigDict, Field

from .local_pdf_index import build_index, match_targets, read_targets


PILOT = ("IDX", "MTS", "EVT", "XRO", "BSL", "CHC", "DMP", "GNC", "S32", "TAH")
ACCEPTED = {"MATCHED_HIGH", "MATCHED_MEDIUM"}
SUCCESS_STATUSES = {"EXTRACTED", "EXTRACTED_WITH_FLAGS"}
EXTRACTION_SCHEMA_VERSION = 3
DEFAULT_MODEL = "gpt-5.6-luna"
DEFAULT_RETRY_MODEL = "gpt-5.6-terra"
# Usage-based safety estimates only; they are not invoice quotes.
MODEL_PRICING_USD_PER_MILLION = {
    "gpt-5.6-luna": {"input": 0.50, "cached_input": 0.05, "output": 4.00},
    "gpt-5.6-terra": {"input": 2.00, "cached_input": 0.20, "output": 12.00},
}
CONSERVATIVE_ISSUER_RESERVE_USD = 0.25
HALVES = ("2H26", "1H27", "2H27", "1H28", "2H28", "FY29+")
KEYWORDS = (
    "borrowings", "loans and borrowings", "interest-bearing liabilities",
    "interest bearing", "financing facilities", "debt facilities",
    "bank facilities", "liquidity", "capital management", "maturity",
    "maturities", "contractual maturities", "current borrowings",
    "non-current borrowings", "cash and cash equivalents", "net debt",
    "lease liabilities",
    "facility agreement", "facility agreements", "entered into",
    "financial close", "closed", "commencement", "commenced",
    "effective date", "tenor", "term facility", "revolving facility",
    "expiry", "expires", "refinancing", "refinanced",
)

MATURITY_BASES = Literal[
    "EXACT_DATE", "RELATIVE_TENOR_MONTH_START", "MONTH_START_ASSUMPTION", "RANGE_START_MONTH",
    "QUARTER_START_ASSUMPTION", "HALF_YEAR_START_ASSUMPTION",
    "CALENDAR_YEAR_START_ASSUMPTION", "FINANCIAL_YEAR_START_ASSUMPTION",
    "DISCLOSURE_BUCKET_START_ASSUMPTION", "TENOR_DERIVED", "UNDETERMINED",
]
TENOR_BASES = Literal[
    "DISCLOSED_ORIGINAL_TENOR", "DISCLOSED_REMAINING_TENOR",
    "DERIVED_FROM_EXACT_DATES", "DERIVED_FROM_REFERENCE_DATE",
    "ASSUMED_FROM_SCREENING_DATE", "UNDETERMINED",
]
SCREENING_AMOUNT_BASES = Literal[
    "TRANCHE_DRAWN_AMOUNT", "TRANCHE_LIMIT", "AGREEMENT_AMOUNT_UNALLOCATED", "UNAVAILABLE",
]
AMOUNT_ALLOCATION_STATUSES = Literal[
    "AGREEMENT_LEVEL_ONLY", "TRANCHE_LEVEL_DISCLOSED",
    "PARTIAL_TRANCHE_ALLOCATION", "UNAVAILABLE",
]
REFERENCE_TYPES = Literal[
    "BALANCE_DATE", "REPORTING_DATE", "REFINANCING_DATE", "EFFECTIVE_DATE",
    "COMMENCEMENT_DATE", "FINANCIAL_CLOSE_DATE", "OTHER_DISCLOSED_REFERENCE",
    "UNDETERMINED",
]


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class MoneyValue(StrictModel):
    value_m: float | None
    currency: str | None
    source_page: int | None
    source_quote_or_evidence: str | None


class DebtTranche(StrictModel):
    parent_agreement_id: str
    ticker: str
    tranche_name: str
    tranche_description: str | None
    tranche_instrument_type: str | None
    tranche_currency: str | None
    tranche_limit: MoneyValue | None
    tranche_drawn_amount: MoneyValue | None
    tranche_undrawn_amount: MoneyValue | None
    maturity_description: str | None
    maturity_reference_date: str | None
    maturity_reference_type: REFERENCE_TYPES
    tenor_months: int | None
    tenor_description: str | None
    tenor_basis: TENOR_BASES
    exact_maturity_date: str | None
    derived_maturity_date: str | None
    assumed_earliest_maturity_date: str | None
    screening_maturity_date: str | None
    screening_maturity_is_assumed: bool
    maturity_assumption_basis: MATURITY_BASES
    maturity_assumption_explanation: str | None
    screening_amount_m: float | None
    screening_amount_currency: str | None = None
    screening_amount_basis: SCREENING_AMOUNT_BASES
    amount_allocation_status: AMOUNT_ALLOCATION_STATUSES
    source_page: int | None
    source_quote_or_evidence: str | None
    confidence: float = Field(ge=0, le=1)


class DebtAgreement(StrictModel):
    agreement_id: str
    ticker: str
    agreement_name: str
    agreement_description: str | None
    lender_or_market: str | None
    instrument_type: str | None
    currency: str | None
    secured_or_unsecured: str | None
    agreement_facility_limit: MoneyValue | None
    agreement_drawn_amount: MoneyValue | None
    agreement_undrawn_amount: MoneyValue | None
    committed_or_uncommitted: Literal["COMMITTED", "UNCOMMITTED", "UNCLEAR"]
    refinancing_date: str | None
    refinancing_date_basis: str | None
    effective_date: str | None
    effective_date_basis: str | None
    financial_close_date: str | None
    financial_close_date_basis: str | None
    agreement_source_page: int | None
    agreement_source_quote_or_evidence: str | None
    confidence: float = Field(ge=0, le=1)
    is_current_period: bool
    tranches: list[DebtTranche]


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
    debt_agreements: list[DebtAgreement]
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
Return only facts supported by the supplied page text. Never guess. All monetary values must be converted to millions using the disclosed reporting unit, retain their currency, source page, and short source evidence. Use null for undisclosed fields. Exclude lease liabilities, derivatives not explicitly reported as borrowings, and trade payables from debt and maturities. Keep lease liabilities separately. Distinguish committed undrawn headroom from facility limits. Determine the balance date and financial year primarily from report contents. Preserve exact facilities separately from broad maturity buckets. Do not double count them. Dates must use YYYY-MM-DD when reliably disclosed. Page labels in the supplied text are the source page numbers.
One agreement may contain multiple tranches. Create a separate tranche for every different name, series, currency, instrument type, tenor, maturity, amount, lender group, or security ranking. Never combine Facility A and Facility B or multiple bond/private-placement series. Distinguish agreement-level amounts from tranche-level amounts and never allocate agreement totals across tranches without source evidence. Distinguish committed capacity from uncommitted accordions and exclude uncommitted capacity from debt and liquidity totals. Extract only the current reporting period; do not create records from prior-year comparatives.
Keep reporting, refinancing, effective, commencement, financial-close, maturity-reference, exact maturity, and derived maturity dates separate. A reporting or balance date is not a maturity date. Words such as matures, maturing, due, expires, expiry, and repayable on introduce a direct maturity, never a tenor reference date. A disclosed maturity month must not have an original tenor added to it. Only derive from a tenor when a genuine start/reference date is explicitly disclosed and no direct maturity overrides it. Preserve complete source wording for every agreement and tranche. Put a date in exact_maturity_date only when an exact contractual day is directly disclosed. Extract a remaining-tenor reference date and type where stated. Do not copy maturity information between separate tranches. Use null for unsupported values; deterministic post-processing will derive conservative screening fields."""


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--pdf-root", required=True)
    result.add_argument("--targets", required=True)
    result.add_argument("--output", default="outputs/asx_maturity_screen.xlsx")
    selection = result.add_mutually_exclusive_group()
    selection.add_argument("--pilot", action="store_true")
    selection.add_argument("--ticker", type=str.upper)
    selection.add_argument("--tickers", nargs="+", type=str.upper, metavar="TICKER")
    result.add_argument("--match-only", action="store_true")
    result.add_argument("--force", action="store_true")
    result.add_argument("--model", default=DEFAULT_MODEL)
    result.add_argument("--retry-model", default=DEFAULT_RETRY_MODEL)
    result.add_argument("--max-api-cost-usd", type=float)
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


MONTHS = {
    name.upper(): number for number, name in enumerate(
        ("January", "February", "March", "April", "May", "June", "July",
         "August", "September", "October", "November", "December"), 1,
    )
}
MONTHS.update({name[:3].upper(): number for name, number in list(MONTHS.items())})
MONTH_PATTERN = "|".join(MONTHS)


def add_months(value: date, months: int) -> date:
    month_index = value.year * 12 + value.month - 1 + months
    year, month_zero = divmod(month_index, 12)
    month = month_zero + 1
    next_month = date(year + (month == 12), 1 if month == 12 else month + 1, 1)
    last_day = (next_month - timedelta(days=1)).day
    return date(year, month, min(value.day, last_day))


def months_between(start: date, end: date) -> int:
    months = (end.year - start.year) * 12 + end.month - start.month
    return months - (end.day < start.day)


def _month_start(name: str, year: str) -> date:
    numeric_year = int(year)
    if numeric_year < 100:
        numeric_year += 2000
    return date(numeric_year, MONTHS[name.upper()], 1)


def direct_maturity_month(description: str) -> date | None:
    """Return a directly disclosed maturity/expiry month, never a reference date."""
    match = re.search(
        rf"\b(?:MATURES?|MATURING|DUE|EXPIRES?|EXPIRY|REPAYABLE\s+ON)\b"
        rf"(?:\s+IN|\s+ON|\s*[:,-])?\s+({MONTH_PATTERN})[-\s]+(20\d{{2}}|\d{{2}})\b",
        description, re.IGNORECASE,
    )
    return _month_start(match.group(1), match.group(2)) if match else None


def parse_text_date(text: str) -> date | None:
    match = re.search(rf"\b(\d{{1,2}})\s+({MONTH_PATTERN})\s+(20\d{{2}})\b", text, re.IGNORECASE)
    if match:
        try:
            return date(int(match.group(3)), MONTHS[match.group(2).upper()], int(match.group(1)))
        except ValueError:
            return None
    return None


def _financial_year_start(financial_year: int, balance_date: date) -> date:
    prior_year_end = date(financial_year - 1, balance_date.month, balance_date.day)
    return prior_year_end + timedelta(days=1)


def _relative_bucket_start(description: str, balance_date: date) -> tuple[date | None, str | None]:
    text = description.upper().replace("–", "-").replace("—", "-")
    if re.search(r"\b(WITHIN|LESS THAN|UP TO)\s+(ONE|1)\s+YEAR\b", text):
        return balance_date + timedelta(days=1), "within one year"
    match = re.search(r"\b(ONE|1)\s+(?:TO|THROUGH|-)\s+(TWO|2)\s+YEARS?\b", text)
    if match:
        return add_months(balance_date, 12) + timedelta(days=1), "one to two years"
    match = re.search(r"\b(TWO|2)\s+(?:TO|THROUGH|-)\s+(FIVE|5)\s+YEARS?\b", text)
    if match:
        return add_months(balance_date, 24) + timedelta(days=1), "two to five years"
    return None, None


NUMBER_WORDS = {
    "ZERO": 0, "ONE": 1, "TWO": 2, "THREE": 3, "FOUR": 4, "FIVE": 5,
    "SIX": 6, "SEVEN": 7, "EIGHT": 8, "NINE": 9, "TEN": 10,
}


def parse_tenor_months(text: str | None) -> int | None:
    if not text:
        return None
    normalised = text.upper().replace("-", " ")
    for word, number in NUMBER_WORDS.items():
        normalised = re.sub(rf"\b{word}\b", str(number), normalised)
    years = re.search(r"\b(\d+(?:\.\d+)?)\s*YEARS?\b", normalised)
    months = re.search(r"\b(\d+)\s*MONTHS?\b", normalised)
    if years or months:
        return round(float(years.group(1)) * 12) if years and not months else (
            round(float(years.group(1)) * 12) + int(months.group(1)) if years else int(months.group(1))
        )
    return None


def stable_agreement_id(agreement: DebtAgreement) -> str:
    components = (
        agreement.ticker, agreement.agreement_name, agreement.lender_or_market or "",
        agreement.instrument_type or "", agreement.currency or "",
        agreement.agreement_source_quote_or_evidence or "",
    )
    stable = "|".join(re.sub(r"\s+", " ", value.strip().upper()) for value in components)
    return f"{agreement.ticker}-{hashlib.sha256(stable.encode('utf-8')).hexdigest()[:12]}"


def derive_tranche_screening(
    tranche: DebtTranche, balance_date_value: str | None,
    agreement: DebtAgreement | None = None,
) -> DebtTranche:
    """Apply reference-tenor parsing before conservative date assumptions."""
    balance_date = parse_date(balance_date_value)
    description = tranche.maturity_description or tranche.tranche_description or ""
    exact = parse_date(tranche.exact_maturity_date)
    disclosed_month = direct_maturity_month(description)
    screening: date | None = None
    basis: str = "UNDETERMINED"
    explanation: str | None = None
    assumed = False

    if tranche.tenor_months is None:
        tranche.tenor_months = parse_tenor_months(tranche.tenor_description or description)
    reference = parse_date(tranche.maturity_reference_date)
    if not reference and tranche.tenor_months is not None:
        reference_phrase = re.search(
            r"\b(REPORTING DATE|BALANCE DATE|AS AT|MEASURED FROM|ENTERED INTO ON|ENTERED INTO|FROM|AFTER|FOLLOWING|COMMENCING ON|EFFECTIVE FROM)\b(.{0,45})",
            description, re.IGNORECASE,
        )
        if reference_phrase:
            reference = parse_text_date(reference_phrase.group(2))
            if reference:
                phrase = reference_phrase.group(1).upper()
                tranche.maturity_reference_type = (
                    "BALANCE_DATE" if phrase in {"BALANCE DATE", "AS AT"}
                    else "REPORTING_DATE" if phrase == "REPORTING DATE"
                    else "EFFECTIVE_DATE" if phrase == "EFFECTIVE FROM"
                    else "COMMENCEMENT_DATE" if phrase in {"COMMENCING ON", "ENTERED INTO ON", "ENTERED INTO"}
                    else "OTHER_DISCLOSED_REFERENCE"
                )
                tranche.maturity_reference_date = reference.isoformat()
    if not reference and tranche.maturity_reference_type in {"BALANCE_DATE", "REPORTING_DATE"}:
        reference = balance_date
        tranche.maturity_reference_date = reference.isoformat() if reference else None
    if not reference and agreement:
        reference_fields = {
            "REFINANCING_DATE": agreement.refinancing_date,
            "EFFECTIVE_DATE": agreement.effective_date,
            "FINANCIAL_CLOSE_DATE": agreement.financial_close_date,
        }
        reference = parse_date(reference_fields.get(tranche.maturity_reference_type))
        if reference:
            tranche.maturity_reference_date = reference.isoformat()

    if exact:
        screening, basis = exact, "EXACT_DATE"
        explanation = "Exact maturity date disclosed for this facility."
        tranche.derived_maturity_date = None
    elif disclosed_month:
        screening, basis, assumed = disclosed_month, "MONTH_START_ASSUMPTION", True
        explanation = "First day of the directly disclosed contractual maturity month."
        tranche.exact_maturity_date = None
        tranche.derived_maturity_date = None
    elif reference and tranche.tenor_months is not None:
        derived = add_months(reference, tranche.tenor_months)
        tranche.derived_maturity_date = derived.isoformat()
        screening = date(derived.year, derived.month, 1)
        basis, assumed = "RELATIVE_TENOR_MONTH_START", True
        explanation = "First day of the maturity month derived from the disclosed reference date and tenor."
        if tranche.tenor_basis == "UNDETERMINED":
            tranche.tenor_basis = "DERIVED_FROM_REFERENCE_DATE"
    else:
        range_match = re.search(
            rf"\b({MONTH_PATTERN})\s+(20\d{{2}})\s+(?:TO|THROUGH|UNTIL|-)\s+"
            rf"({MONTH_PATTERN})\s+(20\d{{2}})\b",
            description, re.IGNORECASE,
        )
        if range_match:
            screening = _month_start(range_match.group(1), range_match.group(2))
            basis, assumed = "RANGE_START_MONTH", True
            explanation = (
                "Earliest possible maturity month from the disclosed range; actual "
                "maturities may occur anywhere within the range."
            )
        if not screening:
            month_match = re.search(rf"\b({MONTH_PATTERN})\s+(20\d{{2}})\b", description, re.IGNORECASE)
            if month_match:
                screening = _month_start(month_match.group(1), month_match.group(2))
                basis, assumed = "MONTH_START_ASSUMPTION", True
                explanation = "First day of the disclosed maturity month."
        if not screening:
            quarter_match = re.search(r"\bQ([1-4])\s*(20\d{2})\b|\b(20\d{2})\s*Q([1-4])\b", description, re.IGNORECASE)
            if quarter_match:
                quarter = int(quarter_match.group(1) or quarter_match.group(4))
                year = int(quarter_match.group(2) or quarter_match.group(3))
                screening = date(year, 1 + (quarter - 1) * 3, 1)
                basis, assumed = "QUARTER_START_ASSUMPTION", True
                explanation = "First day of the disclosed calendar quarter."
        if not screening:
            half_match = re.search(r"\b([12])H\s*(20\d{2})\b|\b(20\d{2})\s*([12])H\b", description, re.IGNORECASE)
            if half_match:
                half = int(half_match.group(1) or half_match.group(4))
                year = int(half_match.group(2) or half_match.group(3))
                screening = date(year, 1 if half == 1 else 7, 1)
                basis, assumed = "HALF_YEAR_START_ASSUMPTION", True
                explanation = "First day of the disclosed calendar half."
        if not screening:
            financial_year_match = re.search(r"\bFY\s*(20\d{2}|\d{2})\b", description, re.IGNORECASE)
            if financial_year_match and balance_date:
                financial_year = int(financial_year_match.group(1))
                if financial_year < 100:
                    financial_year += 2000
                screening = _financial_year_start(financial_year, balance_date)
                basis, assumed = "FINANCIAL_YEAR_START_ASSUMPTION", True
                explanation = "Start of the disclosed financial year based on the entity balance date."
        if not screening:
            years = set(re.findall(r"\b20\d{2}\b", description))
            if len(years) == 1:
                screening = date(int(next(iter(years))), 1, 1)
                basis, assumed = "CALENDAR_YEAR_START_ASSUMPTION", True
                explanation = "First day of the only disclosed calendar maturity year."
        if not screening and balance_date:
            screening, bucket = _relative_bucket_start(description, balance_date)
            if screening:
                basis, assumed = "DISCLOSURE_BUCKET_START_ASSUMPTION", True
                explanation = f"Earliest date in the disclosed relative bucket ({bucket})."
    if reference and exact and tranche.tenor_months is None:
        tranche.tenor_months = months_between(reference, exact)
        tranche.tenor_basis = "DERIVED_FROM_EXACT_DATES"

    drawn, limit = money(tranche.tranche_drawn_amount), money(tranche.tranche_limit)
    if drawn is not None:
        tranche.screening_amount_m = drawn
        tranche.screening_amount_currency = tranche.tranche_drawn_amount.currency
        tranche.screening_amount_basis = "TRANCHE_DRAWN_AMOUNT"
    elif limit is not None:
        tranche.screening_amount_m = limit
        tranche.screening_amount_currency = tranche.tranche_limit.currency
        tranche.screening_amount_basis = "TRANCHE_LIMIT"
    elif agreement and any(money(value) is not None for value in (
        agreement.agreement_facility_limit, agreement.agreement_drawn_amount,
    )):
        tranche.screening_amount_m = None
        tranche.screening_amount_currency = None
        tranche.screening_amount_basis = "AGREEMENT_AMOUNT_UNALLOCATED"
        tranche.amount_allocation_status = "AGREEMENT_LEVEL_ONLY"
    else:
        tranche.screening_amount_m = None
        tranche.screening_amount_currency = None
        tranche.screening_amount_basis = "UNAVAILABLE"
    amount_count = sum(value is not None for value in (drawn, limit, money(tranche.tranche_undrawn_amount)))
    if tranche.amount_allocation_status == "UNAVAILABLE":
        tranche.amount_allocation_status = (
            "TRANCHE_LEVEL_DISCLOSED" if drawn is not None and limit is not None
            else "PARTIAL_TRANCHE_ALLOCATION" if amount_count
            else "UNAVAILABLE"
        )
    tranche.exact_maturity_date = exact.isoformat() if exact else None
    tranche.assumed_earliest_maturity_date = screening.isoformat() if screening and assumed else None
    tranche.screening_maturity_date = screening.isoformat() if screening else None
    tranche.screening_maturity_is_assumed = assumed
    tranche.maturity_assumption_basis = basis
    tranche.maturity_assumption_explanation = explanation
    return tranche


def _without_comparatives(text: str) -> str:
    """Remove clearly labelled comparative fragments before structural parsing."""
    text = re.sub(r"\(\s*(?:19|20)\d{2}\s*:[^)]*\)", " ", text, flags=re.I)
    text = re.sub(r"(?:prior|previous)\s+year|comparative|corresponding\s+period", " ", text, flags=re.I)
    return re.sub(r"(?:19|20)\d{2}\s*:\s*[^.;\n]+", " ", text)


def named_tenors(text: str) -> list[dict]:
    """Return unambiguous current-period name/tenor pairs from complete page text."""
    cleaned = _without_comparatives(text)
    number = r"(?:\d+(?:\.\d+)?|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)"
    tenor = rf"{number}\s*(?:-\s*)?years?(?:\s*,?\s*{number}\s*months?)?|{number}\s*months?"
    name = r"(?:Facility|Tranche|Series|Bond|Note|Class)\s+[A-Z0-9][A-Z0-9._-]*"
    matches: list[tuple[int, str, str, str]] = []
    patterns = (
        re.compile(rf"(?P<tenor>{tenor})\s*\(\s*(?P<name>{name})\s*\)", re.I),
        re.compile(rf"(?P<name>{name})\s*[:\-]\s*(?P<tenor>{tenor})", re.I),
    )
    for pattern in patterns:
        for match in pattern.finditer(cleaned):
            matches.append((match.start(), match.group("name").strip(), match.group("tenor").strip(), match.group(0)))
    unique = {}
    for _, item_name, tenor_text, evidence in sorted(matches):
        months = parse_tenor_months(tenor_text)
        if months:
            unique[item_name.upper()] = {
                "name": item_name, "tenor_text": tenor_text, "tenor_months": months, "evidence": evidence,
            }
    return list(unique.values())


def split_multi_tenor_tranches(
    agreement: DebtAgreement, raw_page_text: str = "", allow_create: bool = True,
) -> list[DebtTranche]:
    """Split a collapsed record only when evidence names each tenor/tranche pair."""
    if len(agreement.tranches) != 1:
        return agreement.tranches
    original = agreement.tranches[0]
    pairs = named_tenors(raw_page_text or original.source_quote_or_evidence or "")
    pair_names = {pair["name"].upper() for pair in pairs}
    linked = original.tranche_name.upper() in pair_names or any(
        name in (original.source_quote_or_evidence or "").upper() for name in pair_names
    )
    if len(pairs) < 2 or not linked or not allow_create:
        return agreement.tranches
    result = []
    for pair in pairs:
        tranche = original.model_copy(deep=True)
        tranche.tranche_name = pair["name"]
        tranche.tenor_description = pair["tenor_text"]
        tranche.tenor_months = pair["tenor_months"]
        tranche.tenor_basis = "DISCLOSED_REMAINING_TENOR"
        tranche.source_quote_or_evidence = pair["evidence"]
        tranche.tranche_limit = tranche.tranche_drawn_amount = tranche.tranche_undrawn_amount = None
        tranche.screening_amount_m = None
        tranche.screening_amount_currency = None
        tranche.screening_amount_basis = "AGREEMENT_AMOUNT_UNALLOCATED"
        tranche.amount_allocation_status = "AGREEMENT_LEVEL_ONLY"
        result.append(tranche)
    return result


def _safeguard_agreement_amounts(agreement: DebtAgreement, tranche: DebtTranche) -> None:
    """Do not mistake a repeated group/parent amount for a supported child allocation."""
    evidence = (tranche.source_quote_or_evidence or "").lower()
    name = tranche.tranche_name.lower()
    parent_values = {money(agreement.agreement_facility_limit), money(agreement.agreement_drawn_amount)} - {None}
    child_values = (tranche.tranche_limit, tranche.tranche_drawn_amount, tranche.tranche_undrawn_amount)
    repeated = any(money(value) in parent_values for value in child_values if value is not None)
    explicitly_named = name in evidence and bool(re.search(r"\b(?:limit|drawn|balance|amount|facility)\b", evidence))
    group_wording = bool(re.search(r"\b(?:total|combined|facilities|group|borrowings)\b", evidence))
    if repeated and (not explicitly_named or group_wording):
        tranche.tranche_limit = tranche.tranche_drawn_amount = tranche.tranche_undrawn_amount = None
        tranche.amount_allocation_status = "AGREEMENT_LEVEL_ONLY"


def postprocess_agreements(payload: ExtractionPayload, raw_page_text: str = "", allow_split: bool = True) -> None:
    current = []
    for agreement in payload.debt_agreements:
        if not agreement.is_current_period:
            payload.validation_flags.append("PRIOR_PERIOD_DATA_USED")
            continue
        agreement.agreement_id = stable_agreement_id(agreement)
        limit = agreement.agreement_facility_limit
        drawn = agreement.agreement_drawn_amount
        if agreement.agreement_undrawn_amount is None and money(limit) is not None and money(drawn) is not None:
            agreement.agreement_undrawn_amount = MoneyValue(
                value_m=round(money(limit) - money(drawn), 6), currency=(limit.currency or drawn.currency),
                source_page=limit.source_page or drawn.source_page,
                source_quote_or_evidence="Derived as disclosed agreement limit less disclosed agreement drawn amount.",
            )
        agreement.tranches = split_multi_tenor_tranches(agreement, raw_page_text, allow_split)
        for tranche in agreement.tranches:
            tranche.parent_agreement_id = agreement.agreement_id
            _safeguard_agreement_amounts(agreement, tranche)
        agreement.tranches = [
            derive_tranche_screening(tranche, payload.balance_date, agreement)
            for tranche in agreement.tranches
        ]
        current.append(agreement)
    payload.debt_agreements = current


def agreement_structure_flags(payload: ExtractionPayload, raw_page_text: str = "") -> list[str]:
    flags: list[str] = []
    ids: dict[str, DebtAgreement] = {}
    for agreement in payload.debt_agreements:
        raw_pairs = named_tenors(raw_page_text)
        returned = {tranche.tranche_name.upper() for tranche in agreement.tranches}
        linked = bool(returned & {pair["name"].upper() for pair in raw_pairs})
        if linked and len(raw_pairs) > 1 and any(pair["name"].upper() not in returned for pair in raw_pairs):
            flags.extend(("MULTIPLE_TENORS_NOT_SPLIT", "MULTIPLE_TRANCHES_COLLAPSED"))
        if not agreement.tranches and agreement.committed_or_uncommitted != "UNCOMMITTED":
            flags.append("MULTIPLE_TRANCHES_COLLAPSED")
        previous = ids.get(agreement.agreement_id)
        if previous and previous.model_dump() != agreement.model_dump():
            flags.append("AGREEMENT_AMOUNT_DUPLICATION_RISK")
        ids[agreement.agreement_id] = agreement
        for tranche in agreement.tranches:
            evidence = tranche.source_quote_or_evidence or ""
            tenor_count = len(re.findall(r"(?:\d+(?:\.\d+)?|[A-Za-z]+)\s+YEARS?", evidence, re.IGNORECASE))
            if tenor_count > 1 and len(agreement.tranches) == 1:
                flags.extend(("MULTIPLE_TENORS_NOT_SPLIT", "MULTIPLE_TRANCHES_COLLAPSED"))
    return list(dict.fromkeys(flags))


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
    model: str, api_key: str, corrective_feedback: str | None = None,
) -> tuple[ExtractionPayload, dict]:
    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    page_content = "\n\n".join(
        f"--- PDF SOURCE PAGE {page} ---\n{text}" for page, text in zip(source_pages, page_texts)
    )
    response = client.responses.parse(
        model=model,
        input=[
            {"role": "system", "content": EXTRACTION_PROMPT},
            {"role": "user", "content": (
                f"Ticker: {ticker}\nCompany: {company_name}\nMatched legal entity: {matched_legal_entity or ''}\n"
                f"Filename financial year (supporting only): {filename_financial_year or ''}\n"
                f"Corrective validation feedback: {corrective_feedback or 'None'}\n\n{page_content}"
            )},
        ],
        text_format=ExtractionPayload,
    )
    parsed = response.output_parsed
    if parsed is None:
        raise ValueError("Model returned no parsed extraction")
    usage = response.usage
    details = getattr(usage, "input_tokens_details", None)
    return parsed, {
        "input_tokens": getattr(usage, "input_tokens", None),
        "cached_input_tokens": getattr(details, "cached_tokens", 0) or 0,
        "output_tokens": getattr(usage, "output_tokens", None),
        "model": model,
    }


def estimated_api_cost(model: str, input_tokens: int = 0, cached_tokens: int = 0, output_tokens: int = 0) -> float:
    """Usage-based estimate; actual billing may differ."""
    price = MODEL_PRICING_USD_PER_MILLION.get(model, MODEL_PRICING_USD_PER_MILLION[DEFAULT_RETRY_MODEL])
    noncached = max(0, input_tokens - cached_tokens)
    return round((noncached * price["input"] + cached_tokens * price["cached_input"] + output_tokens * price["output"]) / 1_000_000, 6)


def validate_extraction(payload: ExtractionPayload, filename_year: str | None) -> list[str]:
    flags = list(dict.fromkeys(payload.validation_flags + agreement_structure_flags(payload)))
    gross = money(payload.gross_debt_excluding_leases)
    current = money(payload.current_borrowings_excluding_leases)
    non_current = money(payload.non_current_borrowings_excluding_leases)
    if gross is not None and current is not None and non_current is not None:
        if abs(current + non_current - gross) / max(abs(gross), 1e-9) > 0.02:
            flags.append("CURRENT_NON_CURRENT_RECON_FAIL")
    agreements = {agreement.agreement_id: agreement for agreement in payload.debt_agreements}
    drawn = []
    for agreement in agreements.values():
        if agreement.committed_or_uncommitted != "COMMITTED":
            continue
        agreement_drawn = money(agreement.agreement_drawn_amount)
        if agreement_drawn is not None:
            drawn.append(agreement_drawn)
        else:
            tranche_drawn = [money(tranche.tranche_drawn_amount) for tranche in agreement.tranches]
            if any(value is not None for value in tranche_drawn):
                drawn.append(sum(value for value in tranche_drawn if value is not None))
    drawn_total = sum(value for value in drawn if value is not None)
    if gross is not None and any(value is not None for value in drawn):
        if abs(drawn_total - gross) / max(abs(gross), 1e-9) > 0.07:
            flags.append("FACILITY_RECON_FAIL")
    money_objects = [
        payload.gross_debt_excluding_leases, payload.current_borrowings_excluding_leases,
        payload.non_current_borrowings_excluding_leases, payload.undrawn_committed_headroom,
        payload.cash_and_cash_equivalents, payload.net_debt_excluding_leases,
        payload.lease_liabilities,
    ]
    monetary_values = [money(value) for value in money_objects]
    for agreement in agreements.values():
        money_objects.extend((agreement.agreement_facility_limit, agreement.agreement_drawn_amount, agreement.agreement_undrawn_amount))
        monetary_values.extend((money(agreement.agreement_facility_limit), money(agreement.agreement_drawn_amount), money(agreement.agreement_undrawn_amount)))
        limit, undrawn = money(agreement.agreement_facility_limit), money(agreement.agreement_undrawn_amount)
        if limit is not None and undrawn is not None and undrawn > limit:
            flags.append("UNDRAWN_EXCEEDS_LIMIT")
        if agreement.committed_or_uncommitted == "UNCLEAR":
            flags.append("COMMITTED_UNCOMMITTED_CLASSIFICATION_UNCLEAR")
        for tranche in agreement.tranches:
            money_objects.extend((tranche.tranche_limit, tranche.tranche_drawn_amount, tranche.tranche_undrawn_amount))
            monetary_values.extend((money(tranche.tranche_limit), money(tranche.tranche_drawn_amount), money(tranche.tranche_undrawn_amount)))
            if tranche.amount_allocation_status == "PARTIAL_TRANCHE_ALLOCATION":
                flags.append("PARTIAL_TRANCHE_ALLOCATION")
            if tranche.screening_maturity_date and tranche.screening_amount_m is None:
                flags.extend(("TRANCHE_AMOUNT_UNDISCLOSED", "MATURITY_WITHOUT_AMOUNT"))
            if tranche.tenor_months is not None and not tranche.maturity_reference_date and not tranche.exact_maturity_date:
                flags.append("TENOR_WITHOUT_REFERENCE_DATE")
            if tranche.exact_maturity_date and tranche.maturity_reference_date == tranche.exact_maturity_date:
                flags.append("REFERENCE_DATE_USED_AS_MATURITY")
            disclosed_month = direct_maturity_month(tranche.maturity_description or "")
            reference_date = parse_date(tranche.maturity_reference_date)
            if disclosed_month and reference_date and (disclosed_month.year, disclosed_month.month) == (reference_date.year, reference_date.month):
                flags.append("REFERENCE_DATE_USED_AS_MATURITY")
    for bucket in payload.disclosure_buckets:
        money_objects.append(bucket.amount)
        monetary_values.append(money(bucket.amount))
    if any(value is not None and value < 0 for value in monetary_values):
        flags.append("NEGATIVE_DEBT_OR_FACILITY_AMOUNT")
    if gross is not None and not 0 <= gross <= 25_000:
        flags.append("GROSS_DEBT_OUT_OF_RANGE")
    balance_date = parse_date(payload.balance_date)
    if balance_date:
        for agreement in agreements.values():
          for tranche in agreement.tranches:
            maturity = parse_date(tranche.screening_maturity_date)
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
    tranches = [tranche for agreement in agreements.values() for tranche in agreement.tranches]
    if tranches and any(
        tranche.screening_amount_m is not None and not tranche.screening_maturity_date
        for tranche in tranches
    ):
        flags.append("MATURITY_PROFILE_INCOMPLETE")
    if any(tranche.maturity_assumption_basis == "RANGE_START_MONTH" for tranche in tranches):
        flags.append("MATURITY_RANGE_EARLIEST_DATE_ASSUMPTION")
    return list(dict.fromkeys(flags))


def maturity_grid(payload: ExtractionPayload) -> tuple[dict[str, float | None], str]:
    exact = {half: None for half in HALVES}
    inferred = {half: None for half in HALVES}
    facility_allocation_found = False
    exact_found = False
    balance_date = parse_date(payload.balance_date)
    for agreement in payload.debt_agreements:
        if agreement.committed_or_uncommitted == "UNCOMMITTED":
            continue
        for tranche in agreement.tranches:
            screening_date = parse_date(tranche.screening_maturity_date)
            if screening_date and balance_date and screening_date < balance_date:
                continue
            amount = tranche.screening_amount_m
            allocation = allocate_exact(amount, tranche.screening_maturity_date)
            for half, value in allocation.items():
                if value is not None:
                    target = inferred if tranche.screening_maturity_is_assumed else exact
                    target[half] = (target[half] or 0) + value
                    facility_allocation_found = True
                    exact_found = exact_found or not tranche.screening_maturity_is_assumed
    if facility_allocation_found:
        quality = "FACILITY_DATED" if exact_found else "BUCKET_INFERRED"
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
        handle.flush()
        os.fsync(handle.fileno())


def document_fingerprint(selected_pdf: str) -> dict:
    stat = Path(selected_pdf).stat()
    return {"selected_pdf": selected_pdf, "file_size": stat.st_size, "modified_time": int(stat.st_mtime)}


def build_record(
    register: dict, payload: ExtractionPayload | None, status: str, flags: list[str],
    grid: dict[str, float | None], profile_quality: str | None, run_log: dict,
    error: str = "",
) -> dict:
    return {
        "extraction_schema_version": EXTRACTION_SCHEMA_VERSION,
        "ticker": register["ticker"], "target_name": register["target_name"],
        "document_match": register, "document_fingerprint": document_fingerprint(register["selected_pdf"]) if register.get("selected_pdf") else None,
        "selected_pages": payload.source_pages if payload else [],
        "extraction": payload.model_dump(mode="json") if payload else None,
        "status": status, "validation_flags": flags, "maturity_grid": grid,
        "profile_quality": profile_quality, "run_log": run_log, "error": error,
    }


def nearest_screening_maturity(extraction: dict) -> dict:
    dated = []
    balance_date = parse_date(extraction.get("balance_date"))
    for agreement in extraction.get("debt_agreements", []):
        if agreement.get("committed_or_uncommitted") == "UNCOMMITTED":
            continue
        for tranche in agreement.get("tranches", []):
            maturity = parse_date(tranche.get("screening_maturity_date"))
            if maturity and (not balance_date or maturity >= balance_date):
                dated.append((agreement, tranche))
    if not dated:
        return {}
    agreement, nearest = min(dated, key=lambda pair: parse_date(pair[1]["screening_maturity_date"]))
    return {
        "Nearest tranche maturity date": nearest.get("screening_maturity_date"),
        "Nearest tranche name": nearest.get("tranche_name"),
        "Nearest maturity agreement": agreement.get("agreement_name"),
        "Nearest maturity assumed": nearest.get("screening_maturity_is_assumed"),
        "Nearest maturity basis": nearest.get("maturity_assumption_basis"),
        "Agreement facility limit at nearest maturity": (agreement.get("agreement_facility_limit") or {}).get("value_m"),
        "Agreement drawn amount at nearest maturity": (agreement.get("agreement_drawn_amount") or {}).get("value_m"),
        "Nearest tranche limit": (nearest.get("tranche_limit") or {}).get("value_m"),
        "Nearest tranche drawn amount": (nearest.get("tranche_drawn_amount") or {}).get("value_m"),
        "Amount potentially maturing": nearest.get("screening_amount_m"),
        "Screening amount basis": nearest.get("screening_amount_basis"),
        "Amount allocation status": nearest.get("amount_allocation_status"),
        "Upcoming maturity flag": "Y",
        "Upcoming maturity confidence": nearest.get("confidence"),
    }


def committed_agreement_totals(extraction: dict) -> tuple[float | None, float | None]:
    unique = {
        agreement.get("agreement_id"): agreement
        for agreement in extraction.get("debt_agreements", [])
        if agreement.get("committed_or_uncommitted") == "COMMITTED"
    }
    limits = [(agreement.get("agreement_facility_limit") or {}).get("value_m") for agreement in unique.values()]
    undrawn = [(agreement.get("agreement_undrawn_amount") or {}).get("value_m") for agreement in unique.values()]
    return (
        sum(value for value in limits if value is not None) if any(value is not None for value in limits) else None,
        sum(value for value in undrawn if value is not None) if any(value is not None for value in undrawn) else None,
    )


def summary_row(record: dict) -> dict:
    match = record["document_match"]
    extraction = record.get("extraction") or {}
    def amount(field: str) -> float | None:
        value = extraction.get(field)
        return value.get("value_m") if value else None
    committed_limit, committed_undrawn = committed_agreement_totals(extraction)
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
        "Committed facility limit": committed_limit,
        "Committed undrawn headroom": committed_undrawn if committed_undrawn is not None else amount("undrawn_committed_headroom"),
        "Cash": amount("cash_and_cash_equivalents"), "Net debt": amount("net_debt_excluding_leases"),
        **record.get("maturity_grid", {}),
        "2H27 maturity total": record.get("maturity_grid", {}).get("2H27 total"),
        "2H27 maturity flag": "Y" if record.get("maturity_grid", {}).get("2H27 total") not in (None, 0) else "",
        "Profile quality": record.get("profile_quality"), "Selected PDF": match.get("selected_pdf"),
        "Source pages": ", ".join(map(str, record.get("selected_pages", []))),
        "Validation flags": " | ".join(record.get("validation_flags", [])),
        "Extraction notes": extraction.get("extraction_notes") or record.get("error", ""),
        **nearest_screening_maturity(extraction),
    }
    return row


SUMMARY_HEADERS = [
    "Ticker", "Company name", "Matched legal entity", "Match status", "Match confidence",
    "Extraction status", "Extraction confidence", "Model used", "Financial year", "Report type",
    "Balance date", "Reporting currency", "Gross debt ex leases", "Current borrowings",
    "Non-current borrowings", "Committed facility limit", "Committed undrawn headroom", "Cash", "Net debt",
] + [
    "Nearest tranche maturity date", "Nearest tranche name", "Nearest maturity agreement",
    "Nearest maturity assumed", "Nearest maturity basis",
    "Agreement facility limit at nearest maturity", "Agreement drawn amount at nearest maturity",
    "Nearest tranche limit", "Nearest tranche drawn amount", "Amount potentially maturing",
    "Screening amount basis", "Amount allocation status", "Upcoming maturity flag",
    "Upcoming maturity confidence",
] + [f"{half} {kind}" for half in HALVES for kind in ("exact", "inferred", "total")] + [
    "2H27 maturity total", "2H27 maturity flag", "Profile quality", "Selected PDF",
    "Source pages", "Validation flags", "Extraction notes",
]
FACILITY_HEADERS = [
    "ticker", "company_name", "agreement_id", "agreement_name", "tranche_name",
    "tranche_description", "instrument_type", "lender_or_market", "currency",
    "secured_or_unsecured", "committed_or_uncommitted", "refinancing_date",
    "refinancing_date_basis", "effective_date", "effective_date_basis",
    "financial_close_date", "financial_close_date_basis", "maturity_reference_date",
    "maturity_reference_type", "tenor_months", "tenor_description", "tenor_basis",
    "exact_maturity_date", "derived_maturity_date", "assumed_earliest_maturity_date",
    "screening_maturity_date", "screening_maturity_is_assumed",
    "maturity_assumption_basis", "maturity_assumption_explanation",
    "agreement_facility_limit_m", "agreement_facility_limit_currency",
    "agreement_drawn_amount_m", "agreement_drawn_amount_currency",
    "agreement_undrawn_amount_m", "agreement_undrawn_amount_currency",
    "tranche_limit_m", "tranche_limit_currency", "tranche_drawn_amount_m", "tranche_drawn_amount_currency",
    "tranche_undrawn_amount_m", "tranche_undrawn_amount_currency",
    "screening_amount_m", "screening_amount_currency", "screening_amount_basis", "amount_allocation_status",
    "source_page", "source_quote_or_evidence", "confidence", "selected_pdf",
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
    "error_message", "input_tokens", "cached_input_tokens", "output_tokens",
    "estimated_api_cost_usd", "cumulative_estimated_api_cost_usd", "attempts",
]


def detail_rows(records: list[dict]) -> tuple[list[dict], list[dict]]:
    facilities: list[dict] = []
    buckets: list[dict] = []
    for record in records:
        extraction = record.get("extraction") or {}
        for agreement in extraction.get("debt_agreements", []):
            agreement_context = {
                "agreement_id": agreement.get("agreement_id"),
                "agreement_name": agreement.get("agreement_name"),
                "instrument_type": agreement.get("instrument_type"),
                "lender_or_market": agreement.get("lender_or_market"),
                "currency": agreement.get("currency"),
                "secured_or_unsecured": agreement.get("secured_or_unsecured"),
                "committed_or_uncommitted": agreement.get("committed_or_uncommitted"),
                "refinancing_date": agreement.get("refinancing_date"),
                "refinancing_date_basis": agreement.get("refinancing_date_basis"),
                "effective_date": agreement.get("effective_date"),
                "effective_date_basis": agreement.get("effective_date_basis"),
                "financial_close_date": agreement.get("financial_close_date"),
                "financial_close_date_basis": agreement.get("financial_close_date_basis"),
            }
            for field in ("agreement_facility_limit", "agreement_drawn_amount", "agreement_undrawn_amount"):
                value = agreement.get(field)
                agreement_context[f"{field}_m"] = value.get("value_m") if value else None
                agreement_context[f"{field}_currency"] = value.get("currency") if value else None
            tranches = agreement.get("tranches", [])
            if not tranches:
                facilities.append({
                    "ticker": agreement.get("ticker"), "company_name": record["target_name"],
                    "selected_pdf": record["document_match"].get("selected_pdf"), **agreement_context,
                })
            for tranche in tranches:
                row = {
                    "company_name": record["target_name"],
                    "selected_pdf": record["document_match"].get("selected_pdf"),
                    **agreement_context, **tranche,
                }
                row["instrument_type"] = tranche.get("tranche_instrument_type") or agreement.get("instrument_type")
                row["currency"] = tranche.get("tranche_currency") or agreement.get("currency")
                for field in ("tranche_limit", "tranche_drawn_amount", "tranche_undrawn_amount"):
                    value = row.pop(field, None)
                    row[f"{field}_m"] = value.get("value_m") if value else None
                    row[f"{field}_currency"] = value.get("currency") if value else None
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
            values = []
            for header in headers:
                value = row.get(header)
                parsed_value = parse_date(value) if header.lower().endswith("date") else None
                values.append(parsed_value or value)
            sheet.append(values)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in range(1, len(headers) + 1):
            sheet.column_dimensions[get_column_letter(column)].width = min(45, max(12, len(headers[column - 1]) + 2))
            if headers[column - 1].lower().endswith("date"):
                for row_number in range(2, sheet.max_row + 1):
                    sheet.cell(row_number, column).number_format = "dd/mm/yyyy"
            if headers[column - 1] in {"Selected PDF", "selected_pdf"}:
                for row_number, record in enumerate(rows, 2):
                    selected = record.get(headers[column - 1])
                    if selected:
                        sheet.cell(row_number, column).hyperlink = Path(selected).as_uri()
                        sheet.cell(row_number, column).style = "Hyperlink"
        if name == "Summary":
            inferred_fill = PatternFill("solid", fgColor="D9D9D9")
            for column, header in enumerate(headers, 1):
                if header.endswith(" inferred"):
                    for cells in sheet.iter_cols(min_col=column, max_col=column, min_row=2, max_row=sheet.max_row):
                        for cell in cells:
                            cell.fill = inferred_fill
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)


def checkpoint_workbook(path: Path, records: list[dict], register: list[dict]) -> None:
    """Build beside the last valid workbook, then atomically replace it."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.stem}.{os.getpid()}.tmp.xlsx")
    try:
        workbook(temporary, records, register)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def run_local(args: argparse.Namespace, extractor: Callable = extract_with_openai) -> int:
    output_path = Path(args.output)
    output_dir = output_path.parent
    pdf_root = Path(args.pdf_root).resolve(strict=True)
    targets = read_targets(Path(args.targets))
    target_by_ticker = {target["ticker"]: target for target in targets}
    if args.pilot:
        targets = [target for target in targets if target["ticker"] in PILOT]
    elif args.ticker:
        requested = [args.ticker]
        missing = [ticker for ticker in requested if ticker not in target_by_ticker]
        if missing:
            raise SystemExit(f"Requested ticker absent from targets.csv: {', '.join(missing)}")
        targets = [target_by_ticker[ticker] for ticker in requested]
    elif args.tickers:
        requested = list(dict.fromkeys(args.tickers))
        missing = [ticker for ticker in requested if ticker not in target_by_ticker]
        if missing:
            raise SystemExit(f"Requested ticker(s) absent from targets.csv: {', '.join(missing)}")
        targets = [target_by_ticker[ticker] for ticker in requested]
    indexed = build_index(pdf_root, output_dir / "pdf_index.csv")
    candidates, register = match_targets(targets, indexed)
    write_csv(output_dir / "document_match_candidates.csv", candidates)
    write_csv(output_dir / "document_register.csv", register)
    results_path = output_dir / "results.jsonl"
    latest = load_results(results_path)
    initial_records = [latest[item["ticker"]] for item in register if item["ticker"] in latest]
    checkpoint_workbook(output_path, initial_records, register)
    if args.match_only:
        print(f"Local matching complete: {output_dir / 'document_register.csv'} (workbook checkpoint saved)")
        return 0
    if args.tickers:
        not_high = [item["ticker"] for item in register if item["match_status"] != "MATCHED_HIGH"]
        if not_high:
            print(
                "Multi-ticker extraction stopped: every requested issuer must be MATCHED_HIGH. "
                f"Review document_register.csv for: {', '.join(not_high)}",
                file=sys.stderr,
            )
            return 3
    accepted = [item for item in register if item["match_status"] in ACCEPTED]
    api_key = os.getenv("OPENAI_API_KEY")
    if accepted and not api_key:
        print("OPENAI_API_KEY is required after matching when accepted PDFs need extraction.", file=sys.stderr)
        return 2
    cumulative_cost = 0.0
    total_targets = len(register)
    for item_number, item in enumerate(register, 1):
        if args.max_api_cost_usd is not None and item["match_status"] in ACCEPTED:
            if args.max_api_cost_usd - cumulative_cost < CONSERVATIVE_ISSUER_RESERVE_USD:
                print(f"API budget estimate stop before {item['ticker']}; partial workbook preserved.")
                break
        if item["match_status"] not in ACCEPTED:
            record = build_record(item, None, item["match_status"], [], {}, None, {
                "ticker": item["ticker"], "selected_pdf": "", "extraction_status": item["match_status"],
            }, item["match_reason"])
            append_result(results_path, record)
            latest[item["ticker"]] = record
            checkpoint_workbook(output_path, [latest[x["ticker"]] for x in register if x["ticker"] in latest], register)
            print(f"[{item_number}/{total_targets}] {item['ticker']} {item['match_status']} — workbook checkpoint saved")
            continue
        selected = Path(item["selected_pdf"]).resolve(strict=True)
        if not selected.is_relative_to(pdf_root):
            raise ValueError(f"Selected PDF is outside --pdf-root: {selected}")
        fingerprint = document_fingerprint(str(selected))
        previous = latest.get(item["ticker"])
        if (
            not args.force and previous
            and previous.get("status") in SUCCESS_STATUSES
            and previous.get("document_fingerprint") == fingerprint
            and previous.get("extraction_schema_version") == EXTRACTION_SCHEMA_VERSION
        ):
            checkpoint_workbook(output_path, [latest[x["ticker"]] for x in register if x["ticker"] in latest], register)
            print(f"[{item_number}/{total_targets}] {item['ticker']} RESUMED — workbook checkpoint saved")
            continue
        started = datetime.now(timezone.utc)
        monotonic_start = time.monotonic()
        retries = 0
        payload = None
        usage: dict = {}
        attempts: list[dict] = []
        error = ""
        source_pages: list[int] = []
        try:
            source_pages, page_texts = select_relevant_pages(selected)
            if not source_pages:
                raise ValueError("No debt-related source pages were identified")
            for attempt in range(3):
                retries = attempt
                try:
                    attempt_model = args.model if attempt == 0 else args.retry_model
                    feedback = error or None
                    candidate, candidate_usage = extractor(
                        item["ticker"], item["target_name"], item.get("matched_legal_entity"),
                        item.get("financial_year"), source_pages, page_texts, attempt_model, api_key, feedback,
                    )
                    attempt_usage = dict(candidate_usage or {})
                    attempt_usage["model"] = attempt_model
                    attempt_usage["estimated_cost_usd"] = estimated_api_cost(
                        attempt_model, attempt_usage.get("input_tokens") or 0,
                        attempt_usage.get("cached_input_tokens") or 0, attempt_usage.get("output_tokens") or 0,
                    )
                    attempts.append(attempt_usage)
                    payload, usage = candidate, attempt_usage
                    raw_text = "\n\n".join(page_texts)
                    postprocess_agreements(payload, raw_text, allow_split=attempt == 2)
                    if payload.extraction_confidence < 0.60 and attempt < 2:
                        error = "Extraction confidence below 0.60; retrying"
                        continue
                    structural = agreement_structure_flags(payload, raw_text)
                    if structural and attempt < 2:
                        error = f"Agreement/tranche structure requires retry: {' | '.join(structural)}"
                        continue
                    break
                except Exception as exc:
                    error = f"{type(exc).__name__}: {exc}"
                    if attempt == 2:
                        raise
            if payload:
                payload.source_pages = sorted(set(payload.source_pages) | set(source_pages))
                payload.model_used = attempts[-1]["model"]
                flags = validate_extraction(payload, item.get("financial_year"))
                grid, quality = maturity_grid(payload)
                status = "EXTRACTED_WITH_FLAGS" if flags else "EXTRACTED"
            else:
                raise ValueError("Extraction returned no payload")
        except Exception as exc:
            error = error or f"{type(exc).__name__}: {exc}"
            flags, grid, quality, status = [], {}, None, "EXTRACTION_FAILED"
        finished = datetime.now(timezone.utc)
        attempt_cost = sum(item.get("estimated_cost_usd", 0) for item in attempts)
        cumulative_cost += attempt_cost
        run_log = {
            "ticker": item["ticker"], "selected_pdf": str(selected),
            "start_time": started.isoformat(), "finish_time": finished.isoformat(),
            "elapsed_seconds": round(time.monotonic() - monotonic_start, 3),
            "extraction_status": status, "retry_count": retries,
            "model_used": payload.model_used if payload else (attempts[-1]["model"] if attempts else args.model),
            "source_pages": ", ".join(map(str, source_pages)), "error_message": error,
            "input_tokens": sum(x.get("input_tokens") or 0 for x in attempts),
            "cached_input_tokens": sum(x.get("cached_input_tokens") or 0 for x in attempts),
            "output_tokens": sum(x.get("output_tokens") or 0 for x in attempts),
            "estimated_api_cost_usd": attempt_cost, "cumulative_estimated_api_cost_usd": cumulative_cost,
            "attempts": json.dumps(attempts),
        }
        record = build_record(item, payload, status, flags, grid, quality, run_log, error)
        append_result(results_path, record)
        latest[item["ticker"]] = record
        checkpoint_workbook(output_path, [latest[x["ticker"]] for x in register if x["ticker"] in latest], register)
        print(f"[{item_number}/{total_targets}] {item['ticker']} {status} — workbook checkpoint saved")
    records = [latest[item["ticker"]] for item in register if item["ticker"] in latest]
    checkpoint_workbook(output_path, records, register)
    print(f"Workbook: {output_path}")
    return 0


def main_for_args(argv=None, extractor: Callable = extract_with_openai) -> int:
    return run_local(parser().parse_args(argv), extractor=extractor)


def main() -> int:
    return main_for_args()


if __name__ == "__main__":
    raise SystemExit(main())
