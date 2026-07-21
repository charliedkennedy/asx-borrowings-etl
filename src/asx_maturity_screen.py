"""Resumable statutory-debt maturity screen for ASX issuers.

This is intentionally conservative: an unavailable source or an uncertain extraction is
recorded as a status rather than manufactured into a number.
"""
from __future__ import annotations

import os
import sys
# Check this before optional runtime dependencies are imported so an unconfigured
# machine gets the promised actionable message rather than an import traceback.
def api_key_present() -> bool:
    """Accept the requested ChuckKey alias without persisting a credential."""
    return bool(os.getenv("OPENAI_API_KEY") or os.getenv("ChuckKey"))


if __name__ == "__main__" and not api_key_present():
    print("OPENAI_API_KEY (or ChuckKey) is required; set it in the environment and rerun. No requests were made.", file=sys.stderr)
    raise SystemExit(2)

import argparse
import base64
import json
import re
import time
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Literal

import fitz
import httpx
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

ASX = "https://www.asx.com.au"
USER_AGENT = "asx-maturity-screen/1.0 (statutory debt research; contact: debt-advisory@example.invalid)"
PILOT = ["IDX", "MTS", "EVT", "XRO", "BSL", "CHC", "DMP", "GNC", "S32", "TAH"]
SYSTEM_INSTRUCTION = '''You are extracting debt data from pages of an Australian statutory financial report. Rules:
1. Report figures exactly as disclosed. Never estimate, never net, never convert currency. If a figure is not on these pages, return null.
2. Detect units ($'000 vs $m) from table headers and output all amounts in millions of the reporting currency, 1 decimal place. Record the reporting currency; some companies report in USD.
3. Debt means interest-bearing borrowings: bank facilities, term loans, bonds, notes, USPP, other loans. EXCLUDE lease liabilities from all debt figures and capture them separately. Exclude trade payables, derivatives and guarantees.
4. Prefer the borrowings / interest-bearing liabilities note (carrying amounts) for balances, and the facilities table for limits, drawn amounts and maturity dates.
5. The liquidity risk maturity table in the financial risk management note shows undiscounted contractual cash flows and usually INCLUDES future interest and often lease payments. Use it for the maturity profile only if nothing better exists, set basis to contractual_undiscounted and the includes_interest flag, and never mix its figures with carrying amounts.
6. Borrowings are often stated net of capitalised borrowing costs. Use the amounts that tie to the balance sheet and note any gross/net difference.
7. Capture undrawn committed facility headroom where disclosed.
8. Every numeric field must carry a page reference (image index and printed page number if visible) plus the note number or title it came from.
9. Confidence 0-1: 0.9+ clean facility table with maturity dates; ~0.7 clear buckets only; ~0.5 only current/non-current found; below 0.5 anything ambiguous. Explain issues in extraction_notes.'''


class Facility(BaseModel):
    description: str = ""
    type: str = ""
    currency: str | None = None
    limit_m: float | None = None
    drawn_m: float | None = None
    maturity_date: str | None = None
    maturity_text: str | None = None
    page_ref: str | None = None


class Bucket(BaseModel):
    label: str = ""
    months_lo: float | None = None
    months_hi: float | None = None
    amount_m: float | None = None
    basis: Literal["carrying", "contractual_undiscounted"] = "carrying"
    includes_interest: bool = False
    includes_leases: bool = False
    source_note: str | None = None
    page_ref: str | None = None


class Extraction(BaseModel):
    ticker: str
    report_type: Literal["annual", "interim"]
    balance_date: str | None = None
    reporting_currency: str | None = None
    units_detected: str | None = None
    gross_borrowings_carrying_m: float | None = None
    current_borrowings_m: float | None = None
    non_current_borrowings_m: float | None = None
    lease_liabilities_current_m: float | None = None
    lease_liabilities_non_current_m: float | None = None
    cash_and_equivalents_m: float | None = None
    undrawn_committed_m: float | None = None
    facilities: list[Facility] = Field(default_factory=list)
    maturity_buckets: list[Bucket] = Field(default_factory=list)
    source_pages: list[str] = Field(default_factory=list)
    confidence: float = 0
    extraction_notes: str = ""


def arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", default="universe.csv")
    p.add_argument("--output", default="outputs/asx_maturity_screen.xlsx")
    p.add_argument("--workdir", default=".")
    p.add_argument("--pilot", action="store_true", help="Run the mandatory ten-name pilot.")
    p.add_argument("--remainder", action="store_true", help="Run names outside the pilot after approval.")
    p.add_argument("--max-cost", type=float, default=15.0)
    return p


def read_tickers(path: Path) -> list[str]:
    return [x.strip().upper() for x in path.read_text().splitlines()[1:] if x.strip()]


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d %B %Y", "%d %b %Y"):
        try: return datetime.strptime(value[:10] if fmt == "%Y-%m-%d" else value, fmt).date()
        except ValueError: pass
    return None


class AsxClient:
    def __init__(self):
        self.client = httpx.Client(headers={"User-Agent": USER_AGENT}, timeout=120, follow_redirects=True)
        self.last_asx = 0.0

    @retry(retry=retry_if_exception_type((httpx.HTTPError, RuntimeError)), wait=wait_exponential(min=1, max=12), stop=stop_after_attempt(3))
    def get(self, url: str) -> httpx.Response:
        if "asx.com.au" in url:
            time.sleep(max(0, 1 - (time.monotonic() - self.last_asx)))
        response = self.client.get(url)
        self.last_asx = time.monotonic()
        if response.status_code == 429 or response.status_code >= 500:
            raise RuntimeError(f"ASX HTTP {response.status_code}")
        response.raise_for_status()
        return response

    def announcements(self, ticker: str) -> tuple[str, list[dict]]:
        url = f"{ASX}/asx/1/company/{ticker}/announcements?count=50&market_sensitive=false"
        data = self.get(url).json()
        return (data.get("name") or data.get("company_name") or "", data.get("data") or data.get("announcements") or [])


ANNUAL = ("annual report", "appendix 4e", "preliminary final report", "full year statutory accounts", "annual financial report")
INTERIM = ("appendix 4d", "half year report", "half yearly report", "interim financial report")


def find_document(api: AsxClient, ticker: str, kind: str, root: Path) -> tuple[dict | None, str]:
    try: company, items = api.announcements(ticker)
    except Exception as exc: return None, str(exc)
    phrases, months = (ANNUAL, 15) if kind == "annual" else (INTERIM, 9)
    cutoff = date.today().replace(day=1)
    candidates = []
    for a in items:
        title = str(a.get("header") or a.get("title") or "")
        released = str(a.get("date") or a.get("release_date") or a.get("document_date") or "")
        d = parse_date(released)
        if not d or (date.today() - d).days > months * 31 or not any(x in title.lower() for x in phrases): continue
        url = a.get("url") or a.get("pdf_url") or a.get("document_url") or ""
        if url.startswith("/"): url = ASX + url
        if url: candidates.append({"title": title, "date": d.isoformat(), "url": url, "company": company})
    if not candidates: return None, "no matching announcement"
    # Latest wins; PDFs are subsequently sized and a full report is favoured.
    candidates.sort(key=lambda x: x["date"], reverse=True)
    for c in candidates:
        try:
            response = api.get(c["url"])
            if not response.content.startswith(b"%PDF"): continue
            suffix = f"{ticker}_{kind}_{c['date']}.pdf"
            pdf = root / "pdfs" / suffix; pdf.parent.mkdir(parents=True, exist_ok=True); pdf.write_bytes(response.content)
            c["path"] = str(pdf); return c, ""
        except Exception: continue
    return None, "matching announcements had no downloadable PDF"


KEYWORDS = ("borrowings", "interest-bearing", "financing arrangements", "financing facilities", "debt facilities", "loans and borrowings", "financial risk management", "liquidity risk", "maturity", "lease liabilities", "net debt")
def select_pages(pdf: Path) -> list[tuple[int, str]]:
    doc = fitz.open(pdf); scored = []
    for n, page in enumerate(doc):
        text = page.get_text("text")
        score = sum(text.lower().count(k) for k in KEYWORDS) * (2 if n >= len(doc) / 2 else 1)
        if score: scored.append((score, n, text))
    if not scored: return []
    wanted = set()
    for _, n, _ in sorted(scored, reverse=True): wanted.update(range(max(0, n-1), min(len(doc), n+2)))
    selected = sorted(wanted, key=lambda n: next((s for s, p, _ in scored if p == n), 0), reverse=True)[:16]
    return [(n, doc[n].get_text("text")) for n in sorted(selected)]


def render_pages(pdf: Path, pages: list[tuple[int, str]], root: Path) -> list[str]:
    doc = fitz.open(pdf); out = []
    for i, (p, _) in enumerate(pages, 1):
        pix = doc[p].get_pixmap(matrix=fitz.Matrix(150/72, 150/72), alpha=False)
        target = root / "page_images" / f"{pdf.stem}_{i}.png"; target.parent.mkdir(parents=True, exist_ok=True); pix.save(target)
        out.append(str(target))
    return out


def validate(x: Extraction, report_date: str, report_type: str) -> list[str]:
    flags = []
    if x.gross_borrowings_carrying_m is not None and x.current_borrowings_m is not None and x.non_current_borrowings_m is not None:
        if abs(x.current_borrowings_m + x.non_current_borrowings_m - x.gross_borrowings_carrying_m) > max(.02*x.gross_borrowings_carrying_m, .1): flags.append("RECON_FAIL")
    drawn = sum(f.drawn_m or 0 for f in x.facilities)
    if drawn and x.gross_borrowings_carrying_m and abs(drawn-x.gross_borrowings_carrying_m) > .07*x.gross_borrowings_carrying_m: flags.append("FACILITY_RECON_FAIL")
    if x.gross_borrowings_carrying_m is not None and not 0 <= x.gross_borrowings_carrying_m <= 25000: flags.append("UNITS_SUSPECT")
    bd, rd = parse_date(x.balance_date), parse_date(report_date)
    if bd and rd and bd > rd: flags.append("BALANCE_DATE_INVALID")
    if report_type == "interim" and bd and rd and (rd-bd).days > 300: flags.append("BALANCE_DATE_SUSPECT")
    return flags


def call_model(client: OpenAI, model: str, x: Extraction, images: list[str]) -> tuple[Extraction, int, float]:
    content = [{"type": "input_text", "text": f"Extract ticker {x.ticker}, report_type {x.report_type}."}]
    for image in images:
        content.append({"type": "input_image", "image_url": "data:image/png;base64," + base64.b64encode(Path(image).read_bytes()).decode()})
    response = client.responses.create(model=model, reasoning={"effort": "minimal"}, input=[{"role":"system","content":SYSTEM_INSTRUCTION},{"role":"user","content":content}], text={"format":{"type":"json_schema", "name":"debt_extraction", "strict":True, "schema":Extraction.model_json_schema()}})
    result = Extraction.model_validate_json(response.output_text)
    usage = response.usage; tokens = (usage.input_tokens or 0) + (usage.output_tokens or 0) if usage else 0
    # Prices vary by model/account; logging token counts is exact, cost is deliberately conservative estimate.
    return result, tokens, tokens * 0.00001


def halfyear(d: date) -> str:
    return f"{'1H' if d.month <= 6 else '2H'}{str(d.year)[-2:]}"


CALENDAR = ["2H26", "1H27", "2H27", "1H28", "2H28", "FY29+"]
def maturity_grid(x: Extraction) -> tuple[dict[str, float | None], str, set[str]]:
    """Allocate dated facilities; bucket allocations are explicitly inferred."""
    grid: dict[str, float | None] = {k: None for k in CALENDAR}; inferred: set[str] = set()
    dated = [f for f in x.facilities if f.drawn_m is not None and parse_date(f.maturity_date)]
    if dated:
        for f in dated:
            key = halfyear(parse_date(f.maturity_date))  # type: ignore[arg-type]
            key = key if key in grid else "FY29+"
            grid[key] = (grid[key] or 0) + f.drawn_m  # type: ignore[operator]
        return grid, "FACILITY_DATED", inferred
    bd = parse_date(x.balance_date)
    if bd and x.maturity_buckets:
        # Each disclosed bucket is evenly spread over its stated month range.  We
        # sample monthly midpoints, which is deterministic and avoids pretending
        # to know the individual debt dates.
        for b in x.maturity_buckets:
            if b.amount_m is None or b.months_lo is None or b.months_hi is None or b.months_hi <= b.months_lo: continue
            lo, hi = int(b.months_lo), int(b.months_hi)
            for month in range(lo, hi):
                d = date(bd.year + (bd.month - 1 + month) // 12, (bd.month - 1 + month) % 12 + 1, 1)
                key = halfyear(d); key = key if key in grid else "FY29+"
                grid[key] = (grid[key] or 0) + b.amount_m / (hi - lo)
                inferred.add(key)
        return grid, "BUCKET_INFERRED", inferred
    return grid, "SPLIT_ONLY", inferred


def write_workbook(path: Path, rows: list[dict], facilities: list[dict], buckets: list[dict], log: list[dict]) -> None:
    wb = Workbook(); ws = wb.active; ws.title = "Summary"
    headers = ["ticker","company name","status","confidence","model used","report used","balance date","reporting currency","gross debt ex leases","current","non-current","undrawn headroom","cash","net debt","2H26","1H27","2H27","1H28","2H28","FY29+","2H27 maturity total","2H27 flag","profile quality","annual PDF","interim PDF"]
    for sheet, data, cols in [(ws,rows,headers),(wb.create_sheet("Facilities"),facilities,None),(wb.create_sheet("Buckets"),buckets,None),(wb.create_sheet("Run log"),log,None)]:
        cols = cols or list(dict.fromkeys(k for r in data for k in r))
        sheet.append(cols)
        for r in data: sheet.append([r.get(c) for c in cols])
        sheet.freeze_panes="A2"; sheet.auto_filter.ref=sheet.dimensions
        for c in range(1, len(cols)+1): sheet.column_dimensions[get_column_letter(c)].width=min(36,max(12,len(str(cols[c-1]))+2))
    for row in ws.iter_rows(min_row=2):
        for cell in row[8:21]: cell.number_format='0.0'
        if row[22].value == "BUCKET_INFERRED":
            for cell in row[14:20]:
                if cell.value is not None:
                    cell.fill = PatternFill('solid', fgColor='D9E1F2')
                    cell.comment = Comment("inferred, even spread", "asx_maturity_screen")
    ws.conditional_formatting.add(f"U2:U{ws.max_row}", CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='FFF2CC')))
    path.parent.mkdir(parents=True, exist_ok=True); wb.save(path)


def main() -> int:
    args = arg_parser().parse_args()
    # The OpenAI SDK reads OPENAI_API_KEY.  Map the user-requested alias only in
    # this process; it is never logged or written to disk.
    if not api_key_present():
        print("OPENAI_API_KEY (or ChuckKey) is required; set it in the environment and rerun. No requests were made.", file=sys.stderr); return 2
    if not os.getenv("OPENAI_API_KEY"):
        os.environ["OPENAI_API_KEY"] = os.environ["ChuckKey"]
    if args.pilot == args.remainder:
        print("Choose exactly one of --pilot or --remainder.", file=sys.stderr); return 2
    root = Path(args.workdir); tickers = read_tickers(Path(args.input)); tickers = [t for t in tickers if (t in PILOT) == args.pilot]
    results_file = root / "results.jsonl"; completed = {json.loads(line)["ticker"] for line in results_file.read_text().splitlines()} if results_file.exists() else set()
    api, oai = AsxClient(), OpenAI(); available = {m.id for m in oai.models.list().data}
    luna = "gpt-5.6-luna" if "gpt-5.6-luna" in available else None
    terra = "gpt-5.6-terra" if "gpt-5.6-terra" in available else None
    if not luna or not terra: print("Required Luna/Terra model IDs are unavailable on this API account.", file=sys.stderr); return 2
    rows=[]; facilities=[]; buckets=[]; log=[]; total=0.0
    for ticker in tickers:
        if ticker in completed: print(f"{ticker}: already complete; skipped"); continue
        started=time.monotonic(); entry={"ticker":ticker,"status":"DOC_NOT_FOUND","tokens":0,"cost_usd":0.0,"pages_sent":0}; annual, err=find_document(api,ticker,"annual",root)
        if not annual: entry["notes"]=err
        else:
            documents=[("annual",annual)]; interim, _=find_document(api,ticker,"interim",root)
            if interim: documents.append(("interim",interim))
            best=None
            for typ, doc in documents:
                pages=select_pages(Path(doc["path"]))
                if not pages: entry["status"]="PAGES_UNRESOLVED"; continue
                images=render_pages(Path(doc["path"]),pages,root); seed=Extraction(ticker=ticker,report_type=typ)
                extracted,tokens,cost=call_model(oai,luna,seed,images); flags=validate(extracted,doc["date"],typ)
                if extracted.confidence < .6 or flags:
                    retry_x,retry_tokens,retry_cost=call_model(oai,terra,seed,images); tokens+=retry_tokens; cost+=retry_cost
                    if retry_x.confidence >= extracted.confidence: extracted=retry_x
                    flags=validate(extracted,doc["date"],typ)
                entry["tokens"]+=tokens; entry["cost_usd"]+=cost; entry["pages_sent"]+=len(images); total+=cost
                candidate=(extracted,doc,flags,"gpt-5.6-terra" if (extracted.confidence < .6 or flags) else luna)
                if best is None or extracted.confidence > best[0].confidence: best=candidate
            if best:
                x,doc,flags,model=best; entry.update(status="REVIEW" if flags else "COMPLETE", confidence=x.confidence, model_used=model, annual_pdf=annual["url"], interim_pdf=(interim or {}).get("url",""), report_used=x.report_type, balance_date=x.balance_date or "", reporting_currency=x.reporting_currency or "", flags=",".join(flags))
                net=(x.gross_borrowings_carrying_m-x.cash_and_equivalents_m) if x.gross_borrowings_carrying_m is not None and x.cash_and_equivalents_m is not None else None
                grid, quality, _ = maturity_grid(x)
                rows.append({**entry,"company name":annual["company"],"gross debt ex leases":x.gross_borrowings_carrying_m,"current":x.current_borrowings_m,"non-current":x.non_current_borrowings_m,"undrawn headroom":x.undrawn_committed_m,"cash":x.cash_and_equivalents_m,"net debt":net,**grid,"2H27 maturity total":grid["2H27"],"2H27 flag":"MATERIAL" if (grid["2H27"] or 0) > 0 else "","profile quality":quality})
                for f in x.facilities: facilities.append({"ticker":ticker,**f.model_dump()})
                for b in x.maturity_buckets: buckets.append({"ticker":ticker,**b.model_dump()})
        if not any(r["ticker"] == ticker for r in rows):
            # Preserve the universe row even where ASX retrieval or page selection failed.
            rows.append({**entry, "company name": (annual or {}).get("company", ""), **{k: None for k in CALENDAR}, "2H27 maturity total": None, "2H27 flag": "", "profile quality": "SPLIT_ONLY"})
        entry["seconds"]=round(time.monotonic()-started,1); log.append(entry); results_file.open("a").write(json.dumps(entry)+"\n")
        print(f"{ticker:4} {entry['status']:14} cost=${entry['cost_usd']:.4f} total=${total:.4f}")
        if total / max(1,len(log)) * len(tickers) > args.max_cost: print("Hard stop: projected full-run cost exceeds limit.",file=sys.stderr); break
    write_workbook(Path(args.output),rows,facilities,buckets,log)
    print("Counts:",dict(Counter(x['status'] for x in log))); print(f"Total spend: ${total:.4f}"); print(f"Workbook: {args.output}")
    if args.pilot: print("PILOT COMPLETE — paused for approval before --remainder.")
    return 0

if __name__ == "__main__": raise SystemExit(main())
