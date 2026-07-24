"""Fast, cached local statutory-report indexing and conservative matching."""
from __future__ import annotations

import csv, re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import fitz
try:
    from rapidfuzz.fuzz import ratio
except ImportError:  # keeps --help usable before optional runtime installation
    from difflib import SequenceMatcher
    def ratio(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio() * 100

SUFFIXES = {"LIMITED","LTD","PTY","PROPRIETARY","HOLDINGS","GROUP","AUSTRALIA","AUSTRALIAN","TRUST","REIT","FUND","STAPLED","COMPANY","CORPORATION","THE"}
BORROWING_WORDS = ("BORROWINGS", "FINANCING FACILITIES", "LOANS AND BORROWINGS", "INTEREST-BEARING")

def normalise_name(value: str) -> str:
    value = value.upper().replace("&", " AND ")
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    words = [w for w in value.split() if w not in SUFFIXES]
    return " ".join(words).replace("GRAIN CORP", "GRAINCORP").replace("BLUE SCOPE", "BLUESCOPE").replace("HEALTH CO", "HEALTHCO").replace("LONG WALE", "LONGWALE")

def parse_filename(filename: str) -> tuple[str, str, str]:
    stem = Path(filename).stem
    m = re.match(r"^(?P<company>.+)_FY(?P<fy>\d{4})_ACN(?P<acn>\d{9})$", stem, re.I)
    return (m.group("company"), m.group("fy"), m.group("acn")) if m else (stem, "", "")

def _metadata(path: Path) -> dict:
    company, fy, acn = parse_filename(path.name); stat = path.stat()
    try:
        doc = fitz.open(path); text = "\n".join(p.get_text("text") for p in list(doc)[:3]); pages = len(doc); doc.close()
        error = ""
    except Exception as exc: text, pages, error = "", 0, f"PDF_UNREADABLE: {type(exc).__name__}: {exc}"
    return {"full_path":str(path),"filename":path.name,"filename_company":company,"financial_year":fy,"acn":acn,"file_size_bytes":stat.st_size,"modified_time":int(stat.st_mtime),"page_count":pages,"first_pages_text":text[:12000],"normalised_filename_company":normalise_name(company),"index_error":error}

def build_index(pdf_root: Path, output: Path, workers: int = 8) -> list[dict]:
    output.parent.mkdir(parents=True, exist_ok=True); cached = {}
    if output.exists():
        with output.open(encoding="utf-8", newline="") as f:
            cached = {r["full_path"]:r for r in csv.DictReader(f)}
    paths = [p for p in pdf_root.rglob("*") if p.is_file() and p.suffix.lower()==".pdf"]
    def one(p: Path):
        old=cached.get(str(p)); s=p.stat()
        return old if old and int(old["file_size_bytes"])==s.st_size and int(old["modified_time"])==int(s.st_mtime) else _metadata(p)
    with ThreadPoolExecutor(max_workers=min(8,max(1,workers))) as pool: rows=list(pool.map(one,paths))
    fields=["full_path","filename","filename_company","financial_year","acn","file_size_bytes","modified_time","page_count","first_pages_text","normalised_filename_company","index_error"]
    with output.open("w",encoding="utf-8",newline="") as f: w=csv.DictWriter(f,fields); w.writeheader(); w.writerows(rows)
    return rows

def read_targets(path: Path) -> list[dict]:
    grouped={}
    with path.open(encoding="utf-8-sig",newline="") as f:
        for r in csv.DictReader(f):
            t=r["ticker"].upper(); grouped.setdefault(t,{"ticker":t,"target_name":r["target_name"],"aliases":[],"warning":""})
            grouped[t]["aliases"] += [x.strip() for x in r.get("aliases","").split("|") if x.strip()] + [r["target_name"]]
    return list(grouped.values())

def match_targets(targets: Iterable[dict], index: list[dict]) -> tuple[list[dict],list[dict]]:
    candidates=[]; register=[]
    for target in targets:
        aliases=list(dict.fromkeys(target["aliases"])); norms=[normalise_name(a) for a in aliases]; scored=[]
        for row in index:
            if row["index_error"]: continue
            fn=row["normalised_filename_company"]; score=max(ratio(n,fn) for n in norms)
            if score < 45: continue
            content=row["first_pages_text"].upper(); content_hit=any(normalise_name(a) in normalise_name(content) for a in aliases)
            annual=any(x in content for x in ("ANNUAL REPORT","FINANCIAL REPORT","CONSOLIDATED FINANCIAL"))
            overall=score + (20 if content_hit else 0)+(5 if annual else 0)+(3 if row["financial_year"] in ("2025","2026") else 0)
            scored.append((overall,score,content_hit,row))
        scored.sort(reverse=True,key=lambda x:x[0]); top=scored[:5]
        for rank,(overall,fscore,hit,row) in enumerate(top,1): candidates.append({"ticker":target["ticker"],"target_name":target["target_name"],"candidate_rank":rank,"candidate_path":row["full_path"],"candidate_filename":row["filename"],"filename_match_score":round(fscore,1),"first_page_match_score":100 if hit else 0,"overall_match_score":round(overall,1),"matched_alias":max(aliases,key=lambda a:ratio(normalise_name(a),row["normalised_filename_company"])),"financial_year":row["financial_year"],"acn":row["acn"],"page_count":row["page_count"],"reason":"first-page identity confirmed" if hit else "filename shortlist only"})
        if not top: status, selected, confidence, reason="NO_LOCAL_DOCUMENT",{},0,"no credible filename candidate"
        elif top[0][2] and (len(top)==1 or top[0][0]-top[1][0]>=8): status,selected,confidence,reason="MATCHED_HIGH",top[0][3],top[0][0],"identity confirmed in first pages"
        elif top[0][2]: status,selected,confidence,reason="MATCHED_MEDIUM",top[0][3],top[0][0],"identity likely; legal-name variation"
        else: status,selected,confidence,reason="AMBIGUOUS_MATCH",{},top[0][0],"filename match not confirmed in PDF"
        register.append({"ticker":target["ticker"],"target_name":target["target_name"],"aliases":" | ".join(aliases),"match_status":status,"match_confidence":round(confidence,1),"selected_pdf":selected.get("full_path",""),"selected_filename":selected.get("filename",""),"financial_year":selected.get("financial_year",""),"acn":selected.get("acn",""),"page_count":selected.get("page_count",""),"match_reason":reason,"alternative_candidates":" | ".join(x[3]["filename"] for x in top[1:]),"validation_warning":target.get("warning","")})
    return candidates,register
