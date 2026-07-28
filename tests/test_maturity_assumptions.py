from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from src.asx_maturity_screen import (
    ExtractionPayload,
    Facility,
    MoneyValue,
    apply_facility_screening,
    derive_facility_screening,
    maturity_grid,
    nearest_screening_maturity,
    workbook,
)


def _money(value: float | None) -> MoneyValue | None:
    if value is None:
        return None
    return MoneyValue(
        value_m=value, currency="AUD", source_page=10,
        source_quote_or_evidence="Facility disclosure",
    )


def _facility(
    description: str | None = None, exact: str | None = None,
    drawn: float | None = 100, limit: float | None = 120,
    close: str | None = None, tenor_months: int | None = None,
    tenor_description: str | None = None,
) -> Facility:
    return Facility(
        ticker="IDX", facility_or_instrument_name="Cash advance facility",
        lender_or_market=None, instrument_type="BANK_FACILITY", currency="AUD",
        facility_limit=_money(limit), drawn_amount=_money(drawn), undrawn_amount=None,
        maturity_description=description, exact_maturity_date=exact,
        assumed_earliest_maturity_date=None, screening_maturity_date=None,
        screening_maturity_is_assumed=False, maturity_assumption_basis="UNDETERMINED",
        maturity_assumption_explanation=None, financial_close_date=close,
        tenor_months=tenor_months, tenor_description=tenor_description,
        tenor_basis="DISCLOSED" if tenor_months is not None else "UNDETERMINED",
        screening_amount_m=None, screening_amount_basis="UNAVAILABLE",
        secured_or_unsecured=None, current_or_non_current=None, source_page=10,
        source_quote_or_evidence=description, confidence=0.9,
    )


def _payload(facilities: list[Facility], balance_date: str = "2025-06-30") -> ExtractionPayload:
    return ExtractionPayload(
        ticker="IDX", company_name="Integral Diagnostics",
        matched_legal_entity="Integral Diagnostics Limited", financial_year="2025",
        balance_date=balance_date, reporting_currency="AUD", reporting_unit="$m",
        gross_debt_excluding_leases=_money(100),
        current_borrowings_excluding_leases=_money(20),
        non_current_borrowings_excluding_leases=_money(80),
        undrawn_committed_headroom=_money(20), cash_and_cash_equivalents=_money(10),
        net_debt_excluding_leases=_money(90), lease_liabilities=_money(5),
        facilities=facilities, disclosure_buckets=[], extraction_confidence=0.9,
        extraction_status="EXTRACTED", model_used="mock", source_pages=[10],
        extraction_notes="", validation_flags=[], profile_quality=None,
        leases_apparently_included=False,
    )


def test_exact_month_range_quarter_half_and_calendar_year() -> None:
    cases = [
        (_facility(exact="2027-12-15"), "2027-12-15", False, "EXACT_DATE"),
        (_facility("December 2027"), "2027-12-01", True, "MONTH_START_ASSUMPTION"),
        (_facility("December 2027 through December 2029"), "2027-12-01", True, "RANGE_START_MONTH"),
        (_facility("matures Q3 2028"), "2028-07-01", True, "QUARTER_START_ASSUMPTION"),
        (_facility("matures in 2H 2028"), "2028-07-01", True, "HALF_YEAR_START_ASSUMPTION"),
        (_facility("matures in 2028"), "2028-01-01", True, "CALENDAR_YEAR_START_ASSUMPTION"),
    ]
    for facility, expected_date, assumed, basis in cases:
        result = derive_facility_screening(facility, "2025-06-30")
        assert result.screening_maturity_date == expected_date
        assert result.screening_maturity_is_assumed is assumed
        assert result.maturity_assumption_basis == basis


def test_financial_year_uses_entity_balance_date() -> None:
    june = derive_facility_screening(_facility("matures FY2028"), "2025-06-30")
    december = derive_facility_screening(_facility("matures FY2028"), "2025-12-31")
    assert june.screening_maturity_date == "2027-07-01"
    assert december.screening_maturity_date == "2028-01-01"
    assert june.maturity_assumption_basis == december.maturity_assumption_basis == "FINANCIAL_YEAR_START_ASSUMPTION"


def test_relative_buckets_and_tenor_derived_maturity() -> None:
    within = derive_facility_screening(_facility("within one year"), "2025-06-30")
    one_two = derive_facility_screening(_facility("one to two years"), "2025-06-30")
    two_five = derive_facility_screening(_facility("two to five years"), "2025-06-30")
    assert within.screening_maturity_date == "2025-07-01"
    assert one_two.screening_maturity_date == "2026-07-01"
    assert two_five.screening_maturity_date == "2027-07-01"
    assert {within.maturity_assumption_basis, one_two.maturity_assumption_basis, two_five.maturity_assumption_basis} == {
        "DISCLOSURE_BUCKET_START_ASSUMPTION"
    }
    tenor = derive_facility_screening(
        _facility(close="2025-03-15", tenor_description="3-year revolving facility"),
        "2025-06-30",
    )
    assert tenor.tenor_months == 36
    assert tenor.tenor_basis == "DISCLOSED"
    assert tenor.screening_maturity_date == "2028-03-15"
    assert tenor.maturity_assumption_basis == "TENOR_DERIVED"
    exact_tenor = derive_facility_screening(
        _facility(exact="2028-03-15", close="2025-03-15"), "2025-06-30",
    )
    assert exact_tenor.tenor_months == 36
    assert exact_tenor.tenor_basis == "DERIVED_FROM_DATES"


def test_screening_amount_priority_and_unsupported_null() -> None:
    drawn = derive_facility_screening(_facility("2028", drawn=80, limit=120), "2025-06-30")
    fallback = derive_facility_screening(_facility("2028", drawn=None, limit=120), "2025-06-30")
    missing = derive_facility_screening(_facility("No maturity information", drawn=None, limit=None), "2025-06-30")
    assert (drawn.screening_amount_m, drawn.screening_amount_basis) == (80, "DRAWN_AMOUNT")
    assert (fallback.screening_amount_m, fallback.screening_amount_basis) == (120, "FACILITY_LIMIT")
    assert missing.screening_amount_m is None
    assert missing.screening_amount_basis == "UNAVAILABLE"
    assert missing.screening_maturity_date is None
    assert missing.maturity_assumption_basis == "UNDETERMINED"


def test_idx_range_is_inferred_once_and_nearest_maturity() -> None:
    cash_advance = _facility(
        "The Group has committed facilities of $468.9m, maturing from December 2027 through December 2029.",
        drawn=343.704, limit=450,
    )
    guarantee = _facility("No separate maturity disclosed", drawn=10, limit=20)
    payload = _payload([cash_advance, guarantee])
    apply_facility_screening(payload)
    cash_advance, guarantee = payload.facilities
    assert cash_advance.exact_maturity_date is None
    assert cash_advance.assumed_earliest_maturity_date == "2027-12-01"
    assert cash_advance.screening_maturity_date == "2027-12-01"
    assert cash_advance.maturity_assumption_basis == "RANGE_START_MONTH"
    assert "actual maturities may occur anywhere within the range" in cash_advance.maturity_assumption_explanation
    assert cash_advance.screening_amount_m == 343.704
    assert cash_advance.screening_amount_basis == "DRAWN_AMOUNT"
    assert guarantee.screening_maturity_date is None
    grid, _ = maturity_grid(payload)
    assert grid["2H27 exact"] is None
    assert grid["2H27 inferred"] == 343.704
    assert grid["2H27 total"] == 343.704
    nearest = nearest_screening_maturity(payload.model_dump(mode="json"))
    assert nearest["Nearest screening maturity date"] == "2027-12-01"
    assert nearest["Nearest maturity facility"] == "Cash advance facility"
    assert nearest["Nearest screening maturity assumed"] is True
    assert nearest["Nearest maturity assumption basis"] == "RANGE_START_MONTH"
    assert nearest["Amount potentially maturing"] == 343.704
    assert nearest["Amount basis"] == "DRAWN_AMOUNT"


def test_exact_and_assumed_facilities_are_not_double_counted() -> None:
    exact = _facility(exact="2027-12-15", drawn=50, limit=60)
    assumed = _facility("December 2027 through December 2029", drawn=100, limit=120)
    payload = _payload([exact, assumed])
    apply_facility_screening(payload)
    grid, _ = maturity_grid(payload)
    assert grid["2H27 exact"] == 50
    assert grid["2H27 inferred"] == 100
    assert grid["2H27 total"] == 150


def test_workbook_dates_use_excel_date_format(tmp_path: Path) -> None:
    payload = _payload([_facility("December 2027", drawn=50)])
    apply_facility_screening(payload)
    grid, quality = maturity_grid(payload)
    selected = tmp_path / "IDX report.pdf"
    selected.write_bytes(b"placeholder")
    match = {
        "ticker": "IDX", "target_name": "Integral Diagnostics",
        "match_status": "MATCHED_HIGH", "match_confidence": 100,
        "selected_pdf": str(selected), "matched_legal_entity": "Integral Diagnostics Limited",
        "financial_year": "2025", "report_type": "ANNUAL_REPORT",
    }
    record = {
        "ticker": "IDX", "target_name": "Integral Diagnostics", "document_match": match,
        "extraction": payload.model_dump(mode="json"), "status": "EXTRACTED",
        "validation_flags": [], "maturity_grid": grid, "profile_quality": quality,
        "selected_pages": [10], "run_log": {}, "error": "",
    }
    output = tmp_path / "screen.xlsx"
    workbook(output, [record], [match])
    book = load_workbook(output)
    facilities = book["Facilities & Instruments"]
    headers = [cell.value for cell in facilities[1]]
    column = headers.index("screening_maturity_date") + 1
    cell = facilities.cell(2, column)
    assert cell.value.date() == date(2027, 12, 1)
    assert cell.number_format == "dd/mm/yyyy"
