from datetime import date

from src.asx_maturity_screen import (
    DEFAULT_MODEL,
    DEFAULT_RETRY_MODEL,
    DebtAgreement,
    DebtTranche,
    derive_tranche_screening,
    named_tenors,
    parse_tenor_months,
    parser,
    split_multi_tenor_tranches,
)


def tranche(**changes):
    values = dict(
        parent_agreement_id="", ticker="TST", tranche_name="Facility A",
        tranche_description=None, tranche_instrument_type="REVOLVING", tranche_currency="AUD",
        tranche_limit=None, tranche_drawn_amount=None, tranche_undrawn_amount=None,
        maturity_description=None, maturity_reference_date=None,
        maturity_reference_type="UNDETERMINED", tenor_months=None, tenor_description=None,
        tenor_basis="UNDETERMINED", exact_maturity_date=None, derived_maturity_date=None,
        assumed_earliest_maturity_date=None, screening_maturity_date=None,
        screening_maturity_is_assumed=False, maturity_assumption_basis="UNDETERMINED",
        maturity_assumption_explanation=None, screening_amount_m=None,
        screening_amount_basis="UNAVAILABLE", amount_allocation_status="UNAVAILABLE",
        source_page=12, source_quote_or_evidence="", confidence=.9,
    )
    values.update(changes)
    return DebtTranche(**values)


def agreement(child):
    return DebtAgreement(
        agreement_id="", ticker="TST", agreement_name="Club debt facility",
        agreement_description=None, lender_or_market=None, instrument_type="CLUB",
        currency="AUD", secured_or_unsecured=None, agreement_facility_limit=None,
        agreement_drawn_amount=None, agreement_undrawn_amount=None,
        committed_or_uncommitted="COMMITTED", refinancing_date=None,
        refinancing_date_basis=None, effective_date=None, effective_date_basis=None,
        financial_close_date=None, financial_close_date_basis=None,
        agreement_source_page=12, agreement_source_quote_or_evidence=None,
        confidence=.9, is_current_period=True, tranches=[child],
    )


def test_raw_page_text_recovers_omitted_named_tranche_and_ignores_comparative():
    raw = ("maturity of two years, six months (Facility A) and four years, six months "
           "(Facility B) (2024: one year, eight months).")
    pairs = named_tenors(raw)
    assert [(p["name"], p["tenor_months"]) for p in pairs] == [("Facility A", 30), ("Facility B", 54)]
    rows = split_multi_tenor_tranches(agreement(tranche(source_quote_or_evidence="Facility A only")), raw)
    assert [(row.tranche_name, row.tenor_months) for row in rows] == [("Facility A", 30), ("Facility B", 54)]
    assert all(row.screening_amount_m is None for row in rows)


def test_three_name_first_tenors_are_detected():
    pairs = named_tenors("Facility A: 2 years 6 months; Facility B: 4 years 6 months; Facility C: 30 months")
    assert [p["tenor_months"] for p in pairs] == [30, 54, 30]


def test_direct_maturity_month_overrides_original_tenor():
    row = derive_tranche_screening(tranche(
        maturity_description="3.5 year Revolving Tranche C1 maturing May-27",
        maturity_reference_date="2027-05-01", maturity_reference_type="OTHER_DISCLOSED_REFERENCE",
        tenor_months=42, tenor_basis="DISCLOSED_ORIGINAL_TENOR",
    ), "2025-06-30")
    assert row.exact_maturity_date is None
    assert row.screening_maturity_date == "2027-05-01"
    assert row.derived_maturity_date is None
    assert row.maturity_assumption_basis == "MONTH_START_ASSUMPTION"


def test_expiry_month_and_original_tenor_derivation_have_correct_roles():
    expiry = derive_tranche_screening(tranche(
        maturity_description="facility expires July 2026", tenor_months=24,
        maturity_reference_date="2026-07-01", maturity_reference_type="OTHER_DISCLOSED_REFERENCE",
    ), "2025-06-30")
    assert expiry.screening_maturity_date == "2026-07-01"
    entered = derive_tranche_screening(tranche(
        maturity_description="entered into on 24 December 2024 for nine years",
        tenor_description="nine years", tenor_basis="DISCLOSED_ORIGINAL_TENOR",
    ), "2025-06-30")
    assert entered.derived_maturity_date == "2033-12-24"
    assert entered.screening_maturity_date == "2033-12-01"


def test_decimal_years_are_months_not_decimal_digits():
    assert parse_tenor_months("3.5 years") == 42


def test_default_and_corrective_models_are_explicit():
    args = parser().parse_args(["--pdf-root", "/tmp/pdfs", "--targets", "targets.csv"])
    assert args.model == DEFAULT_MODEL == "gpt-5.6-luna"
    assert args.retry_model == DEFAULT_RETRY_MODEL == "gpt-5.6-terra"
    assert args.request_timeout_seconds == 240.0
