import sys
from pathlib import Path as _P
sys.path.insert(0, str(_P(__file__).resolve().parents[1]))
import argparse
import csv
from pathlib import Path

from src.schemas import MATURITY_BUCKETS
from src.source_discovery import discover_source
from src.pdf_download import download_pdf
from src.pdf_text_extract import extract_pdf_text_by_page
from src.snippet_finder import find_snippet
from src.borrowings_extractor import extract_borrowings_from_text
from src.qa_checks import run_qa
from src.export_outputs import export_all


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return [r for r in csv.DictReader(f)]


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--pdf-cache", required=True)
    p.add_argument("--skip-source-discovery", action="store_true")
    p.add_argument("--source-register", default="")
    p.add_argument("--limit", type=int, default=0, help="Optional ticker limit for testing")
    return p.parse_args()


def _base_exception(ticker, company, issue_type, description, manual="Y", source_url="", local_pdf_path="", relevant_pages="", snippet_path=""):
    return {
        "ticker": ticker,
        "company_name": company,
        "issue_type": issue_type,
        "issue_description": description,
        "recommended_manual_check": manual,
        "source_url": source_url,
        "local_pdf_path": local_pdf_path,
        "relevant_pages": relevant_pages,
        "extracted_snippet_path": snippet_path,
    }


def main():
    args = parse_args()
    output_dir = Path(args.output)
    (output_dir / "snippets").mkdir(parents=True, exist_ok=True)
    (output_dir / "logs").mkdir(parents=True, exist_ok=True)
    tickers = [r for r in read_csv(args.input) if r.get("ticker")]
    if args.limit:
        tickers = tickers[: args.limit]
    print(f"Loaded {len(tickers)} tickers")

    if args.skip_source_discovery:
        source_records = read_csv(args.source_register)
    else:
        source_records = []
        for idx, r in enumerate(tickers, start=1):
            print(f"[{idx}/{len(tickers)}] Discovering source for {r['ticker']} - {r.get('company_name','')}", flush=True)
            source_records.append(discover_source(r["ticker"], r.get("company_name", "")).__dict__)
    source_by_t = {r["ticker"]: r for r in source_records}

    exceptions, borrow_rows = [], []
    for idx, r in enumerate(tickers, start=1):
        ticker = r["ticker"]
        company = r.get("company_name", "")
        s = source_by_t.get(ticker, {"ticker": ticker, "company_name": company, "source_status": "source not located"})
        print(f"[{idx}/{len(tickers)}] Processing {ticker}: {s.get('source_status')}", flush=True)
        row = {
            "Ticker": ticker,
            "Company Common Name": company,
            "Source report date": s.get("report_date", ""),
            "Source currency": "",
            "Total current borrowings": "not disclosed",
            "Total non-current borrowings": "not disclosed",
            **{b: "not disclosed" for b in MATURITY_BUCKETS},
            "Total borrowings": "not disclosed",
            "Source URL": s.get("source_url", ""),
            "Extraction status": "source not located",
            "Manual review flag": "Y",
            "Notes": s.get("notes", ""),
        }

        if s.get("source_status") != "source located" or not s.get("source_url"):
            exceptions.append(_base_exception(ticker, company, "source_not_located", s.get("notes", "No source URL located")))
            borrow_rows.append(row)
            for issue in run_qa(ticker, company, s, row):
                exceptions.append(issue.__dict__)
            continue

        pdf_path, err = download_pdf(s["source_url"], ticker, Path(args.pdf_cache))
        s["local_pdf_path"] = pdf_path
        if err:
            row["Extraction status"] = "extraction unclear, manual review required"
            row["Notes"] = (row["Notes"] + " " + f"PDF download failed: {err}").strip()
            exceptions.append(_base_exception(ticker, company, "pdf_download_failed", err, "Y", s.get("source_url", "")))
            borrow_rows.append(row)
            for issue in run_qa(ticker, company, s, row):
                exceptions.append(issue.__dict__)
            continue

        pages, quality = extract_pdf_text_by_page(pdf_path)
        snippet_path, relevant_pages = find_snippet(ticker, pages, output_dir / "snippets")
        joined_text = "\n".join(pages)
        parsed = extract_borrowings_from_text(joined_text) if joined_text.strip() else {
            "currency": "",
            "current": "not disclosed",
            "non_current": "not disclosed",
            "total": "not disclosed",
            "maturity": {b: "not disclosed" for b in MATURITY_BUCKETS},
            "status": "extraction unclear, manual review required",
            "manual_review": "Y",
            "notes": f"text extraction quality={quality}",
            "report_date": "",
        }
        if quality != "ok":
            parsed["manual_review"] = "Y"
            parsed["notes"] = (parsed.get("notes", "") + f" Text extraction quality={quality}.").strip()

        row.update({
            "Source report date": parsed.get("report_date") or row["Source report date"],
            "Source currency": parsed["currency"],
            "Total current borrowings": parsed["current"],
            "Total non-current borrowings": parsed["non_current"],
            "Total borrowings": parsed["total"],
            "Extraction status": parsed["status"],
            "Manual review flag": parsed["manual_review"],
            "Notes": (row["Notes"] + " " + parsed.get("notes", "")).strip(),
        })
        for k, v in parsed["maturity"].items():
            row[k] = v
        exceptions.append(_base_exception(ticker, company, "snippet_saved", "Snippet saved for review", "N", s.get("source_url", ""), pdf_path, relevant_pages, snippet_path))
        borrow_rows.append(row)
        for issue in run_qa(ticker, company, s, row):
            exceptions.append(issue.__dict__)

    export_all(source_records, borrow_rows, exceptions, output_dir)
    print(f"Wrote outputs to {output_dir}", flush=True)


if __name__ == "__main__":
    main()
