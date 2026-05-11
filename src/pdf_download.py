from pathlib import Path
from urllib.request import Request, urlopen


def download_pdf(url: str, ticker: str, pdf_cache: Path, timeout: int = 30):
    pdf_cache.mkdir(parents=True, exist_ok=True)
    out = pdf_cache / f"{ticker}.pdf"
    if out.exists() and out.stat().st_size > 0:
        return str(out), None
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=timeout) as r:
            data = r.read()
            ctype = (r.headers.get("Content-Type") or "").lower()
        if "pdf" not in ctype and not data.startswith(b"%PDF"):
            return "", "URL did not return PDF content"
        out.write_bytes(data)
        return str(out), None
    except Exception as e:
        return "", str(e)
