import csv
from pathlib import Path

import fitz

from src.local_pdf_index import (
    INDEX_FIELDS,
    INDEX_PARSER_VERSION,
    build_index,
    classify_report_type,
    match_targets,
    normalise_name,
    parse_filename,
    parse_report_filename,
)


def _pdf(path: Path, text: str) -> None:
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), text)
    document.save(path)
    document.close()


def _indexed(filename: str, text: str, pages: int = 100) -> dict:
    parsed = parse_report_filename(filename)
    return {
        "full_path": filename,
        "filename": filename,
        "index_parser_version": INDEX_PARSER_VERSION,
        **parsed,
        "file_size_bytes": 1000,
        "modified_time": 1,
        "page_count": pages,
        "first_pages_text": text,
        "normalised_filename_company": normalise_name(parsed["filename_company"]),
        "index_error": "",
    }


def _target(ticker: str = "ABG", name: str = "Abacus Group") -> dict:
    return {"ticker": ticker, "target_name": name, "aliases": [name], "warning": ""}


def test_ticker_filename_parsing_and_classification() -> None:
    parsed = parse_report_filename("ABG_2025-08-25_FY25_Annual_Report_and_Appendix_4E.pdf")
    assert parsed == {
        "filename_ticker": "ABG",
        "release_date": "2025-08-25",
        "financial_year": "2025",
        "document_description": "FY25 Annual Report and Appendix 4E",
        "report_type": "ANNUAL_REPORT_AND_4E",
        "filename_company": "",
        "acn": "",
    }
    assert classify_report_type("Appendix_4E") == "APPENDIX_4E"
    assert classify_report_type("Half-Year_Report") == "HALF_YEAR_REPORT"
    assert classify_report_type("Appendix_4D") == "APPENDIX_4D"


def test_real_filename_variants_parse_independent_tokens() -> None:
    cases = {
        "IDX_2025-08-26_Appendix_4E_and_FY25_Annual_Report.pdf":
            ("IDX", "2025-08-26", "2025", "ANNUAL_REPORT_AND_4E"),
        "BSL_2025-08-18_FY2025_Results_for_Announcement_to_Market_&_Annual_Report.pdf":
            ("BSL", "2025-08-18", "2025", "ANNUAL_REPORT"),
        "DMP_2025-08-27_FY25_Appendix_4E_Annual_Report.pdf":
            ("DMP", "2025-08-27", "2025", "ANNUAL_REPORT_AND_4E"),
        "GNC_2025-11-13_FY25_Appendix_4E_and_Annual_Report.pdf":
            ("GNC", "2025-11-13", "2025", "ANNUAL_REPORT_AND_4E"),
    }
    for filename, expected in cases.items():
        parsed = parse_report_filename(filename)
        assert (
            parsed["filename_ticker"], parsed["release_date"],
            parsed["financial_year"], parsed["report_type"],
        ) == expected
    assert parse_report_filename(next(iter(cases)))["document_description"] == (
        "Appendix 4E and FY25 Annual Report"
    )


def test_exact_ticker_survives_missing_optional_metadata() -> None:
    no_year = parse_report_filename("IDX_2025-08-26_Annual_Report.pdf")
    company_title = parse_report_filename("IDX_2025-08-26_Integral_Diagnostics_Annual_Report.PdF")
    unusual = parse_report_filename("IDX_2025-08-26_unusually_named_document.pdf")
    assert no_year["filename_ticker"] == company_title["filename_ticker"] == unusual["filename_ticker"] == "IDX"
    assert no_year["financial_year"] == ""
    assert unusual["financial_year"] == ""
    exact = _indexed(
        "IDX_2025-08-26_Integral_Diagnostics_Annual_Report.pdf",
        "Integral Diagnostics Annual Report",
    )
    fuzzy = _indexed(
        "INTEGRAL DIAGNOSTICS LIMITED_FY2026_ACN123456789.pdf",
        "Integral Diagnostics Annual Report",
    )
    _, register = match_targets(
        [_target("IDX", "Integral Diagnostics")], [fuzzy, exact],
    )
    assert register[0]["selected_filename"] == exact["filename"]
    assert register[0]["match_status"] == "MATCHED_HIGH"


def test_actual_filename_population() -> None:
    cases = [
        ("A1N_2026-02-25_Full_Year_Statutory_Accounts_&_Appendix_4E.pdf", "", "ANNUAL_REPORT_AND_4E"),
        ("A2M_2025-08-18_FY25_Annual_Report.pdf", "2025", "ANNUAL_REPORT"),
        ("AAC_2026-06-18_AACo_2026_Annual_Report.pdf", "2026", "ANNUAL_REPORT"),
        ("ABB_2025-08-25_ABB_FY25_Annual_Report_and_Financial_Statements.pdf", "2025", "ANNUAL_REPORT"),
        ("ABC_2024-02-27_2023_Annual_Report_to_shareholders.pdf", "2023", "ANNUAL_REPORT"),
        ("ABG_2025-08-25_FY25_Annual_Report_and_Appendix_4E.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("ACE_2025-08-26_Acusensus_Appendix_4E_and_FY25_Annual_Report.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("ACF_2025-09-29_Annual_Report_to_shareholders.pdf", "", "ANNUAL_REPORT"),
        ("ACL_2025-09-19_2025_Annual_Report.pdf", "2025", "ANNUAL_REPORT"),
        ("ADH_2025-08-27_ADH_FY2025_Annual_Report_and_Appendix_4E.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("AGL_2026-04-24_Annual_Report_-Year_Ended_31_December_2025.pdf", "2025", "ANNUAL_REPORT"),
        ("AGL_2025-08-13_Appendix_4E_and_2025_Annual_Report.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("AJL_2025-08-29_Annual_Report_2024.pdf", "2024", "ANNUAL_REPORT"),
        ("ALD_2026-02-23_2025_Annual_Report.pdf", "2025", "ANNUAL_REPORT"),
        ("ALQ_2026-06-26_Annual_Report_to_Shareholders.pdf", "", "ANNUAL_REPORT"),
        ("AMA_2025-08-22_FY25_Appendix_4E_and_Annual_Report.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("ANN_2025-08-25_Ansell_Full_Year_Statutory_Accounts_and_2025_Annual_Report.pdf", "2025", "ANNUAL_REPORT"),
        ("APE_2026-04-24_Annual_Report_2025.pdf", "2025", "ANNUAL_REPORT"),
        ("AQZ_2025-08-22_4E_&_Annual_Report_to_Shareholders.pdf", "", "ANNUAL_REPORT_AND_4E"),
        ("ARB_2025-08-19_Appendix_4E_and_Annual_Report_FY2025.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("ASK_2025-08-14_FY25_Annual_Report_and_Appendix_4E.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("ATA_2025-08-27_Appendix_4E_and_Annual_Report.pdf", "", "ANNUAL_REPORT_AND_4E"),
        ("AUB_2025-08-26_FY25_Appendix_4E_and_Annual_Report.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
        ("IDX_2025-08-26_Appendix_4E_and_FY25_Annual_Report.pdf", "2025", "ANNUAL_REPORT_AND_4E"),
    ]
    for filename, financial_year, report_type in cases:
        parsed = parse_report_filename(filename)
        assert parsed["filename_ticker"] == filename.split("_", 1)[0]
        assert parsed["release_date"] == filename.split("_", 2)[1]
        assert parsed["financial_year"] == financial_year
        assert parsed["report_type"] == report_type


def test_legacy_filename_and_normalisation() -> None:
    assert parse_filename("A.P. Eagers Ltd_FY2026_ACN123456789.Pdf") == (
        "A.P. Eagers Ltd", "2026", "123456789",
    )
    assert normalise_name("A.P. Eagers & Co. Limited") == "A P EAGERS AND CO"


def test_exact_ticker_outranks_fuzzy_name_and_newest_annual_wins() -> None:
    rows = [
        _indexed("ABG_2024-08-20_FY24_Annual_Report.pdf", "Abacus Group Annual Report"),
        _indexed("ABG_2025-08-25_FY25_Annual_Report_and_Appendix_4E.pdf", "Abacus Group Annual Report"),
        _indexed("ABACUS GROUP LIMITED_FY2026_ACN123456789.pdf", "Abacus Group Annual Report", 300),
    ]
    candidates, register = match_targets([_target()], rows)
    assert register[0]["match_status"] == "MATCHED_HIGH"
    assert register[0]["selected_filename"] == "ABG_2025-08-25_FY25_Annual_Report_and_Appendix_4E.pdf"
    assert candidates[0]["filename_ticker"] == "ABG"
    assert all(candidate["overall_match_score"] <= 100 for candidate in candidates)


def test_presentation_is_not_selected_as_annual_report() -> None:
    row = _indexed("ABG_2025-08-25_FY25_Full_Year_Results_Presentation.PDF", "Abacus Group results presentation")
    _, register = match_targets([_target()], [row])
    assert row["report_type"] == "OTHER"
    assert register[0]["match_status"] == "AMBIGUOUS_MATCH"
    assert register[0]["selected_pdf"] == ""


def test_missing_and_ambiguous_matches() -> None:
    _, missing = match_targets([_target()], [])
    assert missing[0]["match_status"] == "NO_LOCAL_DOCUMENT"
    duplicate = _indexed("ABG_2025-08-25_FY25_Annual_Report.pdf", "Abacus Group Annual Report")
    other = {**duplicate, "full_path": "copy/" + duplicate["filename"]}
    _, ambiguous = match_targets([_target()], [duplicate, other])
    assert ambiguous[0]["match_status"] == "AMBIGUOUS_MATCH"


def test_mixed_case_extensions_and_stale_cache_rebuild(tmp_path: Path) -> None:
    root = tmp_path / "reports"
    root.mkdir()
    names = [
        "ABG_2025-08-25_FY25_Annual_Report.pdf",
        "MTS_2025-06-23_FY25_Annual_Report.PDF",
        "XRO_2025-05-16_FY25_Full_Year_Results_and_Annual_Report.Pdf",
    ]
    for name in names:
        _pdf(root / name, "Annual Report")
    cache = tmp_path / "outputs" / "pdf_index.csv"
    cache.parent.mkdir()
    stale_header = ",".join(INDEX_FIELDS)
    stale_row = {field: "" for field in INDEX_FIELDS}
    stale_row.update({
        "index_parser_version": "2", "full_path": "stale", "filename": "stale.pdf",
        "file_size_bytes": "1", "modified_time": "1",
    })
    cache.write_text(
        stale_header + "\n" + ",".join(stale_row[field] for field in INDEX_FIELDS) + "\n",
        encoding="utf-8",
    )

    rows = build_index(root, cache, workers=2)
    assert len(rows) == 3
    assert {row["filename_ticker"] for row in rows} == {"ABG", "MTS", "XRO"}
    with cache.open(encoding="utf-8", newline="") as handle:
        assert set(INDEX_FIELDS).issubset(csv.DictReader(handle).fieldnames or [])
    assert all(row["index_parser_version"] == INDEX_PARSER_VERSION for row in rows)
