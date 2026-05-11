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
    return p.parse_args()


def main():
    args = parse_args()
    output_dir = Path(args.output)
    (output_dir / "snippets").mkdir(parents=True, exist_ok=True)
    (output_dir / "logs").mkdir(parents=True, exist_ok=True)
    tickers = [r for r in read_csv(args.input) if r.get("ticker")]
    source_records = read_csv(args.source_register) if args.skip_source_discovery else [discover_source(r["ticker"], r["company_name"]).__dict__ for r in tickers]
    source_by_t = {r["ticker"]: r for r in source_records}
    exceptions, borrow_rows = [], []
    for r in tickers:
        s = source_by_t.get(r["ticker"], {"ticker": r["ticker"], "company_name": r["company_name"], "source_status": "source not located"})
        row = {"Ticker": r["ticker"], "Company Common Name": r["company_name"], "Source report date": s.get("report_date", ""), "Source currency": "", "Total current borrowings": "not disclosed", "Total non-current borrowings": "not disclosed", **{b: "not disclosed" for b in MATURITY_BUCKETS}, "Total borrowings": "not disclosed", "Source URL": s.get("source_url", ""), "Extraction status": "source not located", "Manual review flag": "Y", "Notes": s.get("notes", "")}
        if s.get("source_status") == "source located" and s.get("source_url"):
            pdf_path, err = download_pdf(s["source_url"], r["ticker"], Path(args.pdf_cache))
            s["local_pdf_path"] = pdf_path
            if not err:
                pages, quality = extract_pdf_text_by_page(pdf_path)
                snippet_path, relevant_pages = find_snippet(r["ticker"], pages, output_dir / "snippets")
                parsed = extract_borrowings_from_text("\n".join(pages)) if quality == "ok" else {"currency":"","current":"not disclosed","non_current":"not disclosed","total":"not disclosed","maturity":{b:"not disclosed" for b in MATURITY_BUCKETS},"status":"extraction unclear, manual review required","manual_review":"Y","notes":f"text extraction quality={quality}"}
                row.update({"Source currency": parsed["currency"],"Total current borrowings": parsed["current"],"Total non-current borrowings": parsed["non_current"],"Total borrowings": parsed["total"],"Extraction status": parsed["status"],"Manual review flag": parsed["manual_review"],"Notes": (row["Notes"]+" "+parsed["notes"]).strip()})
                for k, v in parsed["maturity"].items(): row[k]=v
                exceptions.append({"ticker":r["ticker"],"company_name":r["company_name"],"issue_type":"snippet_saved","issue_description":"Snippet saved for review","recommended_manual_check":"N","source_url":s.get("source_url",""),"local_pdf_path":pdf_path,"relevant_pages":relevant_pages,"extracted_snippet_path":snippet_path})
            else:
                exceptions.append({"ticker":r["ticker"],"company_name":r["company_name"],"issue_type":"pdf_download_failed","issue_description":err,"recommended_manual_check":"Y","source_url":s.get("source_url",""),"local_pdf_path":"","relevant_pages":"","extracted_snippet_path":""})
        borrow_rows.append(row)
        for issue in run_qa(r["ticker"], r["company_name"], s, row): exceptions.append(issue.__dict__)
    export_all(source_records, borrow_rows, exceptions, output_dir)

if __name__ == "__main__":
    main()
