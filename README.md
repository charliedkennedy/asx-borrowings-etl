# ASX Borrowings ETL

## Statutory maturity screen

`src/asx_maturity_screen.py` is the resumable, evidence-first statutory filing
pipeline. It requires `OPENAI_API_KEY` in the environment, verifies the pinned
Luna and Terra model IDs before work, writes PDFs under `pdfs/`, resumable
records to `results.jsonl`, and creates the requested four-tab workbook.

Run the mandatory pilot first; the command stops after the ten named issuers.
Only run the remainder after the pilot is reviewed and approved.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# `ChuckKey` is also accepted and is mapped only in the process at runtime.
# For GitHub Actions, save it as the repository secret `CHUCKKEY`.
export OPENAI_API_KEY='...'
python src/asx_maturity_screen.py --pilot --input universe.csv --output outputs/asx_maturity_screen.xlsx
# after approval:
python src/asx_maturity_screen.py --remainder --input universe.csv --output outputs/asx_maturity_screen.xlsx
```

## Install
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Input file
Create `tickers.csv` with columns:
- `ticker`
- `company_name`

## Run full process
```bash
python src/main.py --input tickers.csv --output outputs --pdf-cache pdf_cache
```

## Rerun extraction using edited source register
```bash
python src/main.py --input tickers.csv --skip-source-discovery --source-register outputs/source_register.csv --output outputs --pdf-cache pdf_cache
```

## Outputs
- `outputs/source_register.csv`
- `outputs/borrowings_maturity_profile.csv`
- `outputs/borrowings_maturity_profile.xlsx`
- `outputs/exceptions_report.csv`
- `outputs/snippets/`
- `outputs/logs/`
- `pdf_cache/`

## Status labels
- `source not located`
- `report located but borrowings not found`
- `nil borrowings`
- `current/non-current extracted, maturity not disclosed`
- `current/non-current extracted, partial maturity extracted`
- `full maturity profile extracted`
- `extraction unclear, manual review required`

Use manual review flag and exceptions report for verification.

## Local Windows PDF workflow

The default local mode never contacts ASX. It indexes financial-statement PDFs under the supplied directory, writes `outputs\\pdf_index.csv`, ranks candidates, and writes `outputs\\document_match_candidates.csv` and `outputs\\document_register.csv`.

```powershell
python src\asx_maturity_screen.py --pilot --match-only --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output outputs\asx_maturity_screen.xlsx
```

For extraction after reviewing the register, set the credential without echoing it and run the pilot:

```powershell
$secureKey = Read-Host "OpenAI API key" -AsSecureString
$env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new('', $secureKey).Password
python src\asx_maturity_screen.py --pilot --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output outputs\asx_maturity_screen.xlsx
Remove-Item Env:OPENAI_API_KEY
```
