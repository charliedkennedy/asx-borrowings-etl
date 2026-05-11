# ASX Borrowings ETL

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

## Run from GitHub Actions
1. Go to **Actions** in the GitHub repository.
2. Select **Run Borrowings ETL**.
3. Click **Run workflow**.
4. When the run completes, download the **borrowings-etl-outputs** artifact.
