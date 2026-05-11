from src.qa_checks import run_qa


def test_reconciliation_issue_created():
    issues = run_qa("ABC", "A Co", {"source_url":"https://asx.com.au/x.pdf","report_date":"2026-01-01"}, {
        "Total current borrowings": 1.0,
        "Total non-current borrowings": 2.0,
        "Total borrowings": 4.0,
    })
    assert any(i.issue_type == "reconciliation" for i in issues)
