from dataclasses import dataclass, field
from typing import Dict, List, Optional

MATURITY_BUCKETS = [
    "Jun-26","Dec-26","Jun-27","Dec-27","Jun-28","Dec-28","Jun-29","Dec-29",
    "Jun-30","Dec-30","Jun-31","Dec-31","Jun-32","Dec-32","Jun-33","Dec-33",
    "Jun-34","Dec-34","Jun-35","Dec-35",">2035"
]

EXTRACTION_STATUSES = {
    "source not located",
    "report located but borrowings not found",
    "nil borrowings",
    "current/non-current extracted, maturity not disclosed",
    "current/non-current extracted, partial maturity extracted",
    "full maturity profile extracted",
    "extraction unclear, manual review required",
}

@dataclass
class SourceRecord:
    ticker: str
    company_name: str
    source_status: str = "source not located"
    report_title: str = ""
    report_date: str = ""
    reporting_period: str = ""
    source_type: str = ""
    source_url: str = ""
    local_pdf_path: str = ""
    source_confidence: str = "low"
    notes: str = ""

@dataclass
class ExceptionRecord:
    ticker: str
    company_name: str
    issue_type: str
    issue_description: str
    recommended_manual_check: str
    source_url: str = ""
    local_pdf_path: str = ""
    relevant_pages: str = ""
    extracted_snippet_path: str = ""

@dataclass
class BorrowingRecord:
    ticker: str
    company_name: str
    source_report_date: str = ""
    source_currency: str = ""
    total_current_borrowings: str = "not disclosed"
    total_non_current_borrowings: str = "not disclosed"
    maturity: Dict[str, str] = field(default_factory=lambda: {k: "not disclosed" for k in MATURITY_BUCKETS})
    total_borrowings: str = "not disclosed"
    source_url: str = ""
    extraction_status: str = "source not located"
    manual_review_flag: str = "Y"
    notes: str = ""
