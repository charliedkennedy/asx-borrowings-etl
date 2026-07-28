import json
from pathlib import Path

import fitz
from openpyxl import load_workbook

from src.asx_maturity_screen import (
    DisclosureBucket,
    DebtAgreement,
    DebtTranche,
    ExtractionPayload,
    MoneyValue,
    SUCCESS_STATUSES,
    main_for_args,
    maturity_grid,
    summary_row,
    to_millions,
    validate_extraction,
)


def _money(value: float | None, page: int = 1) -> MoneyValue:
    return MoneyValue(
        value_m=value, currency="AUD", source_page=page,
        source_quote_or_evidence="Borrowings table evidence",
    )


def _payload(
    ticker: str = "IDX", facilities: list[DebtTranche] | None = None,
    buckets: list[DisclosureBucket] | None = None,
) -> ExtractionPayload:
    return ExtractionPayload(
        ticker=ticker, company_name="Integral Diagnostics",
        matched_legal_entity="Integral Diagnostics Limited", financial_year="2025",
        balance_date="2025-06-30", reporting_currency="AUD", reporting_unit="$m",
        gross_debt_excluding_leases=_money(100),
        current_borrowings_excluding_leases=_money(40),
        non_current_borrowings_excluding_leases=_money(60),
        undrawn_committed_headroom=_money(20), cash_and_cash_equivalents=_money(10),
        net_debt_excluding_leases=_money(90), lease_liabilities=_money(15),
        debt_agreements=[_agreement(ticker, facilities)] if facilities else [], disclosure_buckets=buckets or [],
        extraction_confidence=0.9, extraction_status="EXTRACTED", model_used="mock-model",
        source_pages=[1], extraction_notes="Lease liabilities were separately excluded.",
        validation_flags=[], profile_quality=None, leases_apparently_included=False,
    )


def _facility(drawn: float = 100) -> DebtTranche:
    return DebtTranche(
        parent_agreement_id="", ticker="IDX", tranche_name="Syndicated facility",
        tranche_description=None, tranche_instrument_type="BANK_FACILITY", tranche_currency="AUD",
        tranche_limit=_money(120), tranche_drawn_amount=_money(drawn), tranche_undrawn_amount=_money(20),
        maturity_description="15 October 2027", exact_maturity_date="2027-10-15",
        derived_maturity_date=None, maturity_reference_date=None, maturity_reference_type="UNDETERMINED",
        assumed_earliest_maturity_date=None, screening_maturity_date="2027-10-15",
        screening_maturity_is_assumed=False, maturity_assumption_basis="EXACT_DATE",
        maturity_assumption_explanation="Exact maturity date disclosed.",
        tenor_months=None, tenor_description=None,
        tenor_basis="UNDETERMINED", screening_amount_m=drawn,
        screening_amount_basis="TRANCHE_DRAWN_AMOUNT", amount_allocation_status="TRANCHE_LEVEL_DISCLOSED",
        source_page=1, source_quote_or_evidence="Facility matures October 2027",
        confidence=0.9,
    )


def _agreement(ticker: str, facilities: list[DebtTranche]) -> DebtAgreement:
    drawn_total = sum(item.tranche_drawn_amount.value_m for item in facilities if item.tranche_drawn_amount)
    return DebtAgreement(
        agreement_id="mock", ticker=ticker, agreement_name="Syndicated agreement",
        agreement_description=None, lender_or_market=None, instrument_type="BANK_FACILITY",
        currency="AUD", secured_or_unsecured="UNSECURED", agreement_facility_limit=_money(120),
        agreement_drawn_amount=_money(drawn_total), agreement_undrawn_amount=_money(120 - drawn_total),
        committed_or_uncommitted="COMMITTED", refinancing_date=None, refinancing_date_basis=None,
        effective_date=None, effective_date_basis=None, financial_close_date=None,
        financial_close_date_basis=None, agreement_source_page=1,
        agreement_source_quote_or_evidence="Syndicated agreement", confidence=0.9,
        is_current_period=True, tranches=facilities,
    )


def _bucket() -> DisclosureBucket:
    return DisclosureBucket(
        ticker="IDX", bucket_label="1 to 2 years", period_start="2027-07-01",
        period_end="2027-12-31", amount=_money(12), currency="AUD", source_page=1,
        source_quote_or_evidence="1 to 2 years: $12m", confidence=0.8,
    )


def _pdf(path: Path, company: str) -> None:
    document = fitz.open()
    page = document.new_page()
    page.insert_text(
        (72, 72),
        f"{company}\nAnnual Report\nBorrowings and financing facilities\n"
        "Current borrowings 40; non-current borrowings 60; maturity October 2027",
    )
    document.save(path)
    document.close()


def _inputs(tmp_path: Path, two_targets: bool = False) -> tuple[Path, Path, Path]:
    root = tmp_path / "local PDF files"
    root.mkdir()
    _pdf(root / "IDX_2025-08-26_FY25_Annual_Report.pdf", "Integral Diagnostics")
    lines = ["ticker,target_name,aliases", "IDX,Integral Diagnostics,Integral Diagnostics Limited"]
    if two_targets:
        _pdf(root / "MTS_2025-06-23_FY25_Annual_Report.pdf", "Metcash Limited")
        lines.append("MTS,Metcash Limited,Metcash")
    targets = tmp_path / "targets.csv"
    targets.write_text("\n".join(lines) + "\n", encoding="utf-8")
    output = tmp_path / "outputs" / "asx_maturity_screen.xlsx"
    return root, targets, output


def test_units_leases_exact_bucket_and_split_profiles() -> None:
    assert to_millions(125_000, "$'000") == 125
    assert to_millions(125_000_000, "units") == 125
    exact, quality = maturity_grid(_payload(facilities=[_facility()], buckets=[_bucket()]))
    assert exact["2H27 exact"] == 100
    assert exact["2H27 inferred"] is None
    assert quality == "FACILITY_DATED"
    inferred, quality = maturity_grid(_payload(buckets=[_bucket()]))
    assert inferred["2H27 inferred"] == 12
    assert quality == "BUCKET_INFERRED"
    split, quality = maturity_grid(_payload())
    assert all(value is None for value in split.values())
    assert quality == "SPLIT_ONLY"
    record = {
        "ticker": "IDX", "target_name": "Integral Diagnostics",
        "document_match": {"match_status": "MATCHED_HIGH", "match_confidence": 100,
                           "report_type": "ANNUAL_REPORT", "selected_pdf": "report.pdf"},
        "extraction": _payload().model_dump(mode="json"), "status": "EXTRACTED",
        "validation_flags": [], "maturity_grid": split, "profile_quality": quality,
        "selected_pages": [1], "error": "",
    }
    assert summary_row(record)["Gross debt ex leases"] == 100
    assert summary_row(record)["Gross debt ex leases"] != money_value(record, "lease_liabilities")


def money_value(record: dict, field: str) -> float | None:
    value = record["extraction"].get(field)
    return value.get("value_m") if value else None


def test_reconciliation_validation_flags() -> None:
    payload = _payload(facilities=[_facility(50)])
    payload.current_borrowings_excluding_leases = _money(20)
    flags = validate_extraction(payload, "2025")
    assert "CURRENT_NON_CURRENT_RECON_FAIL" in flags
    assert "FACILITY_RECON_FAIL" in flags


def test_resumability_force_ticker_and_workbook(monkeypatch, tmp_path: Path, capsys) -> None:
    root, targets, output = _inputs(tmp_path, two_targets=True)
    monkeypatch.setenv("OPENAI_API_KEY", "test-secret-must-not-be-logged")
    calls: list[str] = []

    def extractor(ticker, *args):
        calls.append(ticker)
        payload = _payload(ticker=ticker, facilities=[_facility()])
        payload.ticker = ticker
        return payload, {"input_tokens": 100, "output_tokens": 50}

    command = ["--ticker", "IDX", "--pdf-root", str(root), "--targets", str(targets), "--output", str(output)]
    assert main_for_args(command, extractor=extractor) == 0
    assert calls == ["IDX"]
    assert main_for_args(command, extractor=extractor) == 0
    assert calls == ["IDX"]
    results_path = output.parent / "results.jsonl"
    stale = json.loads(results_path.read_text().splitlines()[-1])
    stale["extraction_schema_version"] = 1
    with results_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(stale) + "\n")
    assert main_for_args(command, extractor=extractor) == 0
    assert calls == ["IDX", "IDX"]
    assert main_for_args([*command, "--force"], extractor=extractor) == 0
    assert calls == ["IDX", "IDX", "IDX"]
    captured = capsys.readouterr()
    assert "test-secret-must-not-be-logged" not in captured.out + captured.err

    book = load_workbook(output)
    assert book.sheetnames == [
        "Summary", "Facilities & Instruments", "Disclosure Buckets",
        "Document Register", "Exceptions", "Run Log",
    ]
    headers = [cell.value for cell in book["Summary"][1]]
    for required in ("Gross debt ex leases", "2H27 exact", "2H27 inferred", "2H27 total", "Source pages"):
        assert required in headers
    facility_headers = [cell.value for cell in book["Facilities & Instruments"][1]]
    for required in ("agreement_id", "tranche_name", "tranche_drawn_amount_m", "screening_maturity_date", "source_page"):
        assert required in facility_headers
    bucket_headers = [cell.value for cell in book["Disclosure Buckets"][1]]
    for required in ("bucket_label", "period_start", "period_end", "amount_m"):
        assert required in bucket_headers
    selected_column = headers.index("Selected PDF") + 1
    assert book["Summary"].cell(2, selected_column).hyperlink is not None
    records = [json.loads(line) for line in results_path.read_text().splitlines()]
    assert records[-1]["status"] in SUCCESS_STATUSES
    assert records[-1]["extraction"]["lease_liabilities"]["value_m"] == 15


def test_ambiguous_and_failed_extractions_are_safe(monkeypatch, tmp_path: Path) -> None:
    root, targets, output = _inputs(tmp_path)
    original = root / "IDX_2025-08-26_FY25_Annual_Report.pdf"
    duplicate_dir = root / "duplicate"
    duplicate_dir.mkdir()
    duplicate = duplicate_dir / original.name
    duplicate.write_bytes(original.read_bytes())
    calls = 0

    def should_not_run(*args):
        nonlocal calls
        calls += 1
        raise AssertionError("ambiguous matches must not be extracted")

    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    assert main_for_args(["--pdf-root", str(root), "--targets", str(targets), "--output", str(output)], extractor=should_not_run) == 0
    assert calls == 0

    duplicate.unlink()
    monkeypatch.setenv("OPENAI_API_KEY", "not-logged")
    attempts = 0

    def failing(*args):
        nonlocal attempts
        attempts += 1
        raise ValueError("mock structured extraction failure")

    assert main_for_args(["--force", "--pdf-root", str(root), "--targets", str(targets), "--output", str(output)], extractor=failing) == 0
    assert attempts == 3
    latest = json.loads((output.parent / "results.jsonl").read_text().splitlines()[-1])
    assert latest["status"] == "EXTRACTION_FAILED"
    assert latest["extraction"] is None
