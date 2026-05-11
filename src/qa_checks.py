from .schemas import ExceptionRecord


def run_qa(ticker, company, source_row, borrow_row):
    issues = []
    c = borrow_row["Total current borrowings"]
    n = borrow_row["Total non-current borrowings"]
    t = borrow_row["Total borrowings"]
    if isinstance(c, (int, float)) and isinstance(n, (int, float)) and isinstance(t, (int, float)):
        if round(c + n, 3) != round(t, 3):
            issues.append(ExceptionRecord(ticker, company, "reconciliation", "current + non-current != total", "Y", source_row.get("source_url", ""), source_row.get("local_pdf_path", "")))
    if not source_row.get("report_date"):
        issues.append(ExceptionRecord(ticker, company, "missing_report_date", "source report date missing", "Y", source_row.get("source_url", ""), source_row.get("local_pdf_path", "")))
    if source_row.get("source_url") and not any(x in source_row.get("source_url", "").lower() for x in ["asx.com.au", ".com.au"]):
        issues.append(ExceptionRecord(ticker, company, "source_officiality", "source URL may be non-official", "Y", source_row.get("source_url", ""), source_row.get("local_pdf_path", "")))
    return issues
