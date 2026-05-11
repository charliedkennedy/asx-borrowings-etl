from pathlib import Path
from typing import List, Tuple

TERMS = [
    "borrowings","loans and borrowings","current borrowings","non-current borrowings",
    "interest-bearing liabilities","bank loans","bank debt","secured loans","unsecured loans",
    "financing facilities","financial liabilities","contractual maturities","maturity analysis",
    "maturity profile","debt facilities","lease liabilities"
]


def find_snippet(ticker: str, pages: List[str], snippet_dir: Path) -> Tuple[str, str]:
    snippet_dir.mkdir(parents=True, exist_ok=True)
    hits = []
    for i, txt in enumerate(pages, start=1):
        low = txt.lower()
        if any(t in low for t in TERMS):
            hits.append((i, txt[:3000]))
    path = snippet_dir / f"{ticker}_borrowings_snippet.txt"
    if not hits:
        path.write_text("No keyword hit found.")
        return str(path), ""
    content = []
    pages_idx = []
    for p, t in hits[:5]:
        pages_idx.append(str(p))
        content.append(f"=== PAGE {p} ===\n{t}\n")
    path.write_text("\n".join(content), encoding="utf-8")
    return str(path), ",".join(pages_idx)
