import csv
from pathlib import Path

import fitz

from src.local_pdf_index import (
    INDEX_FIELDS,
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
        "document_description": "Annual Report and Appendix 4E",
        "report_type": "ANNUAL_REPORT_AND_4E",
        "filename_company": "",
        "acn": "",
    }
    assert classify_report_type("Appendix_4E") == "APPENDIX_4E"
    assert classify_report_type("Half-Year_Report") == "HALF_YEAR_REPORT"
    assert classify_report_type("Appendix_4D") == "APPENDIX_4D"


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
    cache.write_text("full_path,filename,file_size_bytes,modified_time\nstale,stale.pdf,1,1\n", encoding="utf-8")

    rows = build_index(root, cache, workers=2)
    assert len(rows) == 3
    assert {row["filename_ticker"] for row in rows} == {"ABG", "MTS", "XRO"}
    with cache.open(encoding="utf-8", newline="") as handle:
        assert set(INDEX_FIELDS).issubset(csv.DictReader(handle).fieldnames or [])
