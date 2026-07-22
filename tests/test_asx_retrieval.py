import pytest

pytest.importorskip("httpx")

from src.asx_maturity_screen import ASX, announcement_items, parse_date


def test_nested_announcements_and_relative_url():
    item = announcement_items({"data": {"announcements": [{"document_release_date": "2026-02-20T08:30:00+1100", "url": "/document.pdf"}]}})[0]
    assert str(parse_date(item["document_release_date"])) == "2026-02-20"
    assert ASX + item["url"] == "https://www.asx.com.au/document.pdf"


def test_iso_offset_with_colon():
    assert str(parse_date("2026-02-20T08:30:00+11:00")) == "2026-02-20"
