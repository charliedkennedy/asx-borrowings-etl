from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from src.asx_maturity_screen import (
    DebtAgreement, DebtTranche, DisclosureBucket, ExtractionPayload, MoneyValue,
    derive_tranche_screening, maturity_grid, nearest_screening_maturity,
    normalize_extraction, postprocess_agreements, remaining_annual_instalment_dates,
    summary_row, validate_extraction, workbook,
)


def cash(value, currency="AUD", role="UNKNOWN"):
    return MoneyValue(value_m=value, currency=currency, source_page=1,
                      source_quote_or_evidence="fixture evidence", amount_role=role)


def tranche(name="Facility A", **changes):
    values = dict(
        parent_agreement_id="", ticker="TST", tranche_name=name, tranche_description=None,
        tranche_instrument_type="REVOLVING", tranche_currency="AUD", tranche_limit=None,
        tranche_drawn_amount=None, tranche_undrawn_amount=None, maturity_description=None,
        maturity_reference_date=None, maturity_reference_type="UNDETERMINED", tenor_months=None,
        tenor_description=None, tenor_basis="UNDETERMINED", exact_maturity_date=None,
        derived_maturity_date=None, assumed_earliest_maturity_date=None,
        screening_maturity_date=None, screening_maturity_is_assumed=False,
        maturity_assumption_basis="UNDETERMINED", maturity_assumption_explanation=None,
        screening_amount_m=None, screening_amount_basis="UNAVAILABLE",
        amount_allocation_status="UNAVAILABLE", source_page=1,
        source_quote_or_evidence="fixture evidence", confidence=.9,
    )
    values.update(changes)
    return DebtTranche(**values)


def agreement(children, **changes):
    values = dict(
        agreement_id="", ticker="TST", agreement_name="Test facility",
        agreement_description=None, lender_or_market=None, instrument_type="REVOLVING",
        currency="AUD", secured_or_unsecured=None, agreement_facility_limit=None,
        agreement_drawn_amount=None, agreement_undrawn_amount=None,
        committed_or_uncommitted="COMMITTED", refinancing_date=None,
        refinancing_date_basis=None, effective_date=None, effective_date_basis=None,
        financial_close_date=None, financial_close_date_basis=None, agreement_source_page=1,
        agreement_source_quote_or_evidence="Test facility fixture evidence", confidence=.9,
        is_current_period=True, tranches=children,
    )
    values.update(changes)
    return DebtAgreement(**values)


def payload(agreements=None, buckets=None):
    return ExtractionPayload(
        ticker="TST", company_name="Model name ignored", matched_legal_entity="Test Legal Limited",
        financial_year="2025", balance_date="2025-06-30", reporting_currency="AUD",
        reporting_unit="$m", gross_debt_excluding_leases=cash(100),
        current_borrowings_excluding_leases=cash(20), non_current_borrowings_excluding_leases=cash(80),
        undrawn_committed_headroom=None, cash_and_cash_equivalents=cash(120),
        net_debt_excluding_leases=cash(-20), lease_liabilities=None,
        debt_agreements=agreements or [], disclosure_buckets=buckets or [], extraction_confidence=.9,
        extraction_status="EXTRACTED", model_used="mock", source_pages=[1], extraction_notes="",
        validation_flags=[], profile_quality=None, leases_apparently_included=False,
    )


def test_as_of_date_excludes_historical_maturity_but_retains_detail():
    row = tranche(exact_maturity_date="2025-12-15", tranche_drawn_amount=cash(50, role="FACILITY_DRAWN"))
    data = payload([agreement([row])])
    postprocess_agreements(data)
    grid, _ = maturity_grid(data, date(2026, 7, 1))
    assert all(grid[f"{half} total"] is None for half in ("2H26", "1H27", "2H27", "1H28", "2H28", "FY29+"))
    assert nearest_screening_maturity(data.model_dump(mode="json"), date(2026, 7, 1)) == {}
    assert "MATURITY_BEFORE_SCREENING_AS_OF_DATE" in validate_extraction(data, "2025", date(2026, 7, 1))
    assert data.debt_agreements[0].tranches[0].screening_maturity_date == "2025-12-15"


def test_measurement_date_is_not_maturity_and_upper_bound_is_not_earliest():
    measurement = derive_tranche_screening(tranche(
        maturity_description="fully utilised at 30 June 2025",
        maturity_reference_date="2025-06-30", maturity_reference_type="BALANCE_DATE",
    ), "2025-06-30")
    assert measurement.screening_maturity_date is None
    upper = derive_tranche_screening(tranche(
        maturity_description="facility maturities staggered through to September 2027",
    ), "2025-06-30")
    assert upper.screening_maturity_date is None
    assert "latest/final maturity endpoint" in upper.maturity_assumption_explanation


def test_financial_year_period_uses_actual_year_end():
    june = derive_tranche_screening(tranche(maturity_description="Financial Year of Debt Maturity FY2026"), "2025-06-30")
    december = derive_tranche_screening(tranche(maturity_description="Financial Year of Debt Maturity FY2026"), "2025-12-31")
    assert (june.screening_maturity_date, june.maturity_period_end) == ("2025-07-01", "2026-06-30")
    assert (december.screening_maturity_date, december.maturity_period_end) == ("2026-01-01", "2026-12-31")
    assert june.maturity_assumption_basis == "FINANCIAL_YEAR_START_ASSUMPTION"


def test_contractual_buckets_get_boundaries_and_do_not_become_debt_grid():
    buckets = [DisclosureBucket(
        ticker="TST", bucket_label="between one and two years", period_start=None, period_end=None,
        amount=cash(40, role="CONTRACTUAL_CASH_FLOW"), currency="AUD", source_page=1,
        source_quote_or_evidence="contractual cash flows including interest", confidence=.8,
        amount_type="PRINCIPAL_AND_INTEREST", safe_for_summary=False,
    )]
    data = payload(buckets=buckets)
    postprocess_agreements(data)
    assert (buckets[0].period_start, buckets[0].period_end) == ("2026-07-01", "2027-06-30")
    grid, quality = maturity_grid(data)
    assert all(grid[f"{half} total"] is None for half in ("2H26", "1H27", "2H27", "1H28", "2H28", "FY29+"))
    assert quality == "NO_MATURITY_DISCLOSURE"


def test_only_remaining_unambiguous_annual_instalments_are_derived():
    text = ("Four equal annual payments. The second payment was made on 28 February 2025 "
            "and two instalments remaining.")
    assert [item.isoformat() for item in remaining_annual_instalment_dates(text)] == [
        "2026-02-28", "2027-02-28",
    ]
    assert remaining_annual_instalment_dates("annual payments may be made") == []


def test_foreign_currency_and_facility_limit_are_blocked_from_debt_grid_but_capacity_is_separate():
    row = tranche(
        exact_maturity_date="2027-11-15", tranche_limit=cash(400, "AUD", "FACILITY_LIMIT"),
        tranche_drawn_amount=None, exposure_type="FACILITY_CAPACITY",
    )
    foreign = tranche(
        "USD note", exact_maturity_date="2027-11-20", tranche_drawn_amount=cash(125, "USD", "PRINCIPAL_OUTSTANDING"),
        tranche_currency="USD", exposure_type="FUNDED_DEBT",
    )
    data = payload([agreement([row, foreign])])
    postprocess_agreements(data)
    grid, _ = maturity_grid(data)
    assert grid["2H27 total"] is None
    assert grid["2H27 capacity total"] == 400
    assert "REPORTING_CURRENCY_EQUIVALENT_UNAVAILABLE" in validate_extraction(data, "2025")


def test_single_child_amount_propagates_but_multiple_children_stay_unallocated():
    only = tranche(source_quote_or_evidence="Test facility drawn amount and limit")
    one = payload([agreement([only], agreement_drawn_amount=cash(30), agreement_facility_limit=cash(50))])
    postprocess_agreements(one)
    assert one.debt_agreements[0].tranches[0].tranche_drawn_amount.value_m == 30
    many = payload([agreement([tranche(), tranche("Facility B")], agreement_drawn_amount=cash(30))])
    postprocess_agreements(many)
    assert all(item.tranche_drawn_amount is None for item in many.debt_agreements[0].tranches)


def test_cross_currency_or_note_basis_never_calculates_negative_undrawn():
    note = agreement(
        [tranche("USPP series", face_value=cash(300, "USD", "FACE_VALUE"),
                 carrying_amount=cash(457.364, "AUD", "CARRYING_AMOUNT"), tranche_currency="USD")],
        instrument_type="USPP", agreement_facility_limit=cash(300, "USD", "FACE_VALUE"),
        agreement_drawn_amount=cash(457.364, "AUD", "CARRYING_AMOUNT"),
    )
    data = payload([note])
    postprocess_agreements(data)
    assert data.debt_agreements[0].agreement_undrawn_amount is None
    assert "CURRENCY_OR_BASIS_MISMATCH" in validate_extraction(data, "2025")


def test_workbook_has_review_queue_target_company_and_net_debt_fields(tmp_path: Path):
    data = payload()
    match = {"ticker": "TST", "target_name": "Target Company Limited", "match_status": "MATCHED_HIGH",
             "match_confidence": 100, "selected_pdf": "", "report_type": "ANNUAL_REPORT"}
    record = {"ticker": "TST", "target_name": "Target Company Limited", "document_match": match,
              "extraction": data.model_dump(mode="json"), "status": "EXTRACTED", "validation_flags": [],
              "maturity_grid": {}, "profile_quality": "NO_MATURITY_DISCLOSURE", "selected_pages": [1],
              "error": "", "run_log": {}, "screening_as_of_date": "2026-07-29"}
    output = tmp_path / "screen.xlsx"
    workbook(output, [record], [match])
    book = load_workbook(output)
    assert book.sheetnames[-1] == "Review Queue"
    review_headers = [cell.value for cell in book["Review Queue"][1]]
    assert {"agreement_id", "tranche_name", "rejected_raw_value", "normalized_value", "evidence"} <= set(review_headers)
    headers = [cell.value for cell in book["Summary"][1]]
    values = [cell.value for cell in book["Summary"][2]]
    row = dict(zip(headers, values))
    assert row["Company name"] == "Target Company Limited"
    assert row["Calculated net debt ex leases"] == -20
    assert row["Screening as-of date"].strftime("%Y-%m-%d") == "2026-07-29"


def normalize(data, text="AUD financial statements"):
    return normalize_extraction(data, {"ticker": "TST", "target_name": "Canonical Target Limited"}, text, date(2026, 7, 29))


def test_normalization_quarantines_bucket_derived_instrument_values():
    row = tranche(
        maturity_description="between one and five years", tenor_months=60,
        derived_maturity_date="2030-06-30", screening_maturity_date="2030-06-01",
        screening_maturity_is_assumed=True, maturity_assumption_basis="DISCLOSURE_BUCKET_START_ASSUMPTION",
        principal_outstanding=cash(100, role="PRINCIPAL_OUTSTANDING"),
        source_quote_or_evidence="liquidity contractual cash flows between one and five years including interest",
    )
    result = normalize(payload([agreement([row])]))
    item = result.debt_agreements[0].tranches[0]
    assert item.tenor_months is item.derived_maturity_date is item.screening_maturity_date is None
    assert item.debt_maturity_aggregation_eligible is False
    assert "BUCKET_INSTRUMENT_ALLOCATION_UNSUPPORTED" in result.validation_flags
    grid, _ = maturity_grid(result, date(2026, 7, 29))
    assert all(grid[f"{half} total"] is None for half in ("2H26", "1H27", "2H27", "1H28", "2H28", "FY29+"))


def test_normalization_clears_unsupported_exact_date_and_planned_refinancing():
    row = tranche(
        exact_maturity_date="2027-09-30", screening_maturity_date="2027-09-30",
        maturity_description="maturing September 2027", source_quote_or_evidence="maturing September 2027",
    )
    parent = agreement([row], refinancing_date="2027-09-30",
                       agreement_source_quote_or_evidence="facility will refinance and matures September 2027")
    result = normalize(payload([parent]))
    parent, item = result.debt_agreements[0], result.debt_agreements[0].tranches[0]
    assert parent.refinancing_date is None
    assert item.exact_maturity_date is None
    assert item.screening_maturity_date == "2027-09-01"
    assert {audit.validation_code for audit in result.normalization_audit} >= {
        "PLANNED_REFINANCING_NOT_COMPLETED", "EXACT_DATE_EVIDENCE_MISSING",
    }


def test_tenor_range_average_and_original_without_reference_do_not_create_maturity():
    ranged = tranche(maturity_description="committed facilities between 2.2 and 3.2 years",
                     tenor_months=38, screening_maturity_date="2028-08-01")
    average = tranche("Average pool", maturity_description="average maturity of 14 months",
                      tenor_months=14, screening_maturity_date="2026-08-01")
    original = tranche("Original tenor", tenor_description="five years", tenor_months=60,
                       tenor_basis="DISCLOSED_ORIGINAL_TENOR", screening_maturity_date="2030-06-01")
    result = normalize(payload([agreement([ranged, average, original])]))
    first, second, third = result.debt_agreements[0].tranches
    assert (first.tenor_months_min, first.tenor_months_max, first.screening_maturity_date) == (26, 38, None)
    assert second.tenor_is_average and second.screening_maturity_date is None
    assert third.screening_maturity_date is None


def test_maturity_range_retains_period_but_blocks_whole_balance_allocation():
    row = tranche(
        maturity_description="maturities from December 2025 to January 2028",
        source_quote_or_evidence="USD 90m facilities have maturities from December 2025 to January 2028",
        principal_outstanding=cash(90, "USD", "PRINCIPAL_OUTSTANDING"),
        reporting_currency_equivalent_m=135, reporting_currency_equivalent_currency="AUD",
    )
    result = normalize(payload([agreement([row])]))
    item = result.debt_agreements[0].tranches[0]
    assert (item.maturity_period_start, item.maturity_period_end) == ("2025-12-01", "2028-01-31")
    assert item.exact_maturity_date is None and not item.debt_maturity_aggregation_eligible
    grid, _ = maturity_grid(result, date(2026, 7, 29))
    assert all(grid[f"{half} total"] is None for half in ("2H26", "1H27", "2H27", "1H28", "2H28", "FY29+"))


def test_normalized_status_contributions_currency_and_contingent_separation():
    funded = tranche(
        "RCF", exact_maturity_date="2027-12-15", maturity_description="matures 15 December 2027",
        source_quote_or_evidence="AUD 50m drawn under AUD 100m RCF maturing 15 December 2027",
        tranche_drawn_amount=cash(50, role="FACILITY_DRAWN"), tranche_limit=cash(100, role="FACILITY_LIMIT"),
    )
    guarantee = tranche(
        "Guarantee", exact_maturity_date="2027-12-15", maturity_description="expires 15 December 2027",
        source_quote_or_evidence="AUD 20m guarantee utilisation expires 15 December 2027",
        tranche_drawn_amount=cash(20, role="GUARANTEE_UTILISATION"), exposure_type="GUARANTEE_OR_LC",
    )
    foreign = tranche(
        "EUR note", exact_maturity_date="2027-12-15", maturity_description="matures 15 December 2027",
        source_quote_or_evidence="EUR 15m note matures 15 December 2027; AUD equivalent 27m",
        principal_outstanding=cash(15, "EUR", "PRINCIPAL_OUTSTANDING"),
        reporting_currency_equivalent_m=27, reporting_currency_equivalent_currency="AUD",
    )
    past = tranche(
        "Old facility", exact_maturity_date="2025-12-15", maturity_description="matures 15 December 2025",
        source_quote_or_evidence="AUD 10m facility matures 15 December 2025",
        tranche_drawn_amount=cash(10, role="FACILITY_DRAWN"),
    )
    result = normalize(payload([agreement([funded, guarantee, foreign, past])]))
    rows = result.debt_agreements[0].tranches
    assert rows[0].debt_maturity_amount_m == 50 and rows[0].facility_capacity_expiry_amount_m == 100
    assert rows[1].debt_maturity_amount_m is None and rows[1].contingent_expiry_amount_m == 20
    assert rows[2].debt_maturity_amount_reporting_m == 27
    assert rows[3].instrument_status == "PAST_DUE_STATUS_UNKNOWN"
    assert rows[3].debt_maturity_aggregation_eligible is False
    grid, _ = maturity_grid(result, date(2026, 7, 29))
    assert grid["2H27 exact"] == 77
    assert grid["2H27 capacity exact"] == 100
    assert grid["2H27 contingent exact"] == 20


def test_non_iso_currency_and_canonical_identity_are_normalized():
    row = tranche(
        exact_maturity_date="2027-12-15", maturity_description="matures 15 December 2027",
        source_quote_or_evidence="$10m drawn matures 15 December 2027",
        tranche_drawn_amount=cash(10, "$", "FACILITY_DRAWN"),
    )
    result = normalize(payload([agreement([row])]), "AUD financial statements")
    assert result.company_name == "Canonical Target Limited"
    assert result.debt_agreements[0].tranches[0].tranche_drawn_amount.currency == "AUD"
    assert "NON_ISO_CURRENCY_NORMALIZED" in result.validation_flags


def test_idx_style_named_tenors_survive_normalization_without_parent_amount_allocation():
    raw = ("As at 30 June 2025 the bank loan facilities have a maturity of two years, six months "
           "(Facility A) and four years, six months (Facility B) (2024: one year, eight months).")
    child = tranche(
        source_quote_or_evidence="two years, six months (Facility A)",
        maturity_reference_date="2025-06-30", maturity_reference_type="BALANCE_DATE",
        tenor_months=30, tenor_basis="DISCLOSED_REMAINING_TENOR",
    )
    parent = agreement([child], agreement_name="Club debt facility", agreement_drawn_amount=cash(341.252))
    result = normalize(payload([parent]), raw)
    rows = result.debt_agreements[0].tranches
    assert [(row.tranche_name, row.tenor_months) for row in rows] == [("Facility A", 30), ("Facility B", 54)]
    assert [(row.derived_maturity_date, row.screening_maturity_date) for row in rows] == [
        ("2027-12-30", "2027-12-01"), ("2029-12-30", "2029-12-01"),
    ]
    assert all(row.tranche_drawn_amount is None and row.debt_maturity_amount_m is None for row in rows)
    assert all(row.amount_allocation_status == "AGREEMENT_LEVEL_ONLY" for row in rows)


def test_completed_extension_keeps_event_date_separate_from_new_maturity():
    row = tranche(
        exact_maturity_date="2026-08-31", maturity_description="extended to expire on 31 August 2026",
        source_quote_or_evidence="facility was extended on 15 August 2025 to expire on 31 August 2026",
    )
    result = normalize(payload([agreement([row], agreement_source_quote_or_evidence=row.source_quote_or_evidence)]))
    parent, item = result.debt_agreements[0], result.debt_agreements[0].tranches[0]
    assert parent.amendment_or_extension_date == "2025-08-15"
    assert item.exact_maturity_date == "2026-08-31"


def test_shared_limit_counts_only_controller_and_uncommitted_is_excluded():
    controller = tranche(
        "Controlling cap", exact_maturity_date="2027-12-15", maturity_description="matures 15 December 2027",
        source_quote_or_evidence="AUD 50m controlling cap matures 15 December 2027",
        tranche_limit=cash(50, role="FACILITY_LIMIT"), shared_limit_group_id="shared-1",
    )
    sublimit = tranche(
        "LC sublimit", exact_maturity_date="2027-12-15", maturity_description="expires 15 December 2027",
        source_quote_or_evidence="AUD 17m LC sublimit expires 15 December 2027",
        tranche_limit=cash(17, role="FACILITY_LIMIT"), shared_limit_group_id="shared-1",
        is_sublimit=True, exposure_type="GUARANTEE_OR_LC",
    )
    committed = agreement([controller, sublimit])
    optional = agreement(
        [tranche("Accordion", tranche_limit=cash(20, role="OPTIONAL_ACCORDION"))],
        agreement_name="Optional accordion", committed_or_uncommitted="UNCOMMITTED",
    )
    result = normalize(payload([committed, optional]))
    rows = result.debt_agreements[0].tranches
    assert rows[0].facility_capacity_expiry_amount_m == 50
    assert rows[1].facility_capacity_expiry_amount_m is None
    assert result.debt_agreements[1].tranches[0].uncommitted_capacity_amount_m == 20
    grid, _ = maturity_grid(result, date(2026, 7, 29))
    assert grid["2H27 capacity exact"] == 50


def test_holder_put_and_repaid_status_are_normalized_separately_from_final_maturity():
    option = tranche(
        "Convertible", exact_maturity_date="2028-11-30", maturity_description="final maturity 30 November 2028",
        source_quote_or_evidence="holder put 31 May 2027 and final maturity 30 November 2028",
        holder_put_date="2027-05-31", final_contractual_maturity_date="2028-11-30",
        principal_outstanding=cash(25, role="PRINCIPAL_OUTSTANDING"),
    )
    repaid = tranche(
        "Repaid note", exact_maturity_date="2027-12-15", maturity_description="matures 15 December 2027",
        source_quote_or_evidence="the note has been fully repaid; original maturity 15 December 2027",
        principal_outstanding=cash(10, role="PRINCIPAL_OUTSTANDING"),
    )
    result = normalize(payload([agreement([option, repaid])]))
    first, second = result.debt_agreements[0].tranches
    assert first.screening_maturity_date == "2027-05-31"
    assert first.final_contractual_maturity_date == "2028-11-30"
    assert second.instrument_status == "REPAID" and not second.debt_maturity_aggregation_eligible
    assert "OPTION_DATE_BEFORE_FINAL_MATURITY" in result.validation_flags
