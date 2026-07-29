# ASX Borrowings ETL

## Statutory maturity screen

## Local Windows PDF workflow

`src.asx_maturity_screen` is a **local-PDF-only** pipeline. It reads source
documents exclusively from the directory supplied with `--pdf-root`; it has no
web discovery, download, or fallback capability. It indexes financial-statement
PDFs under the supplied directory, writes `outputs\\pdf_index.csv`, ranks
candidates, and writes `outputs\\document_match_candidates.csv` and
`outputs\\document_register.csv`.

```powershell
python -m src.asx_maturity_screen --pilot --match-only --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv
```

For extraction after reviewing the register, set the credential without echoing it and run the pilot:

```powershell
$secureKey = Read-Host "OpenAI API key" -AsSecureString
$env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new('', $secureKey).Password
python -m src.asx_maturity_screen --pilot --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output "outputs\asx_maturity_screen.xlsx"
Remove-Item Env:OPENAI_API_KEY
```

Extraction selects debt-related pages locally, sends only that targeted text to
the configured OpenAI model, appends each completed result to
`outputs\results.jsonl`, and reconstructs the six-sheet workbook. Use
`--ticker IDX` to process one issuer, `--force` to rerun a successful result,
or `--model MODEL_ID` to select an account-accessible structured-output model.
Matching failures and ambiguous matches are never submitted for extraction.
Facility screening distinguishes exact maturity dates from conservative assumed
earliest dates derived from disclosed months, ranges, quarters, halves, years,
relative buckets, or close-date/tenor evidence. Assumed dates populate only the
inferred maturity grid and retain their assumption basis and explanation.
Debt is stored as parent agreements with separate child tranches or instruments.
Agreement totals are repeated only for workbook context and are deduplicated in
issuer totals; unallocated agreement amounts are never copied into tranche
maturity buckets.

## Six-name validation gate

Run matching first without an API key:

```powershell
python -m src.asx_maturity_screen --tickers IDX ORA KPG CMW ALQ WGN --match-only --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output "outputs\asx_maturity_screen.xlsx"
```

Inspect `outputs\document_register.csv`. Extraction is blocked unless every
issuer selected through `--tickers` is `MATCHED_HIGH`. After all six matches are
confirmed, enter the key without echoing it and force a fresh six-name run:

```powershell
$secureKey = Read-Host "OpenAI API key" -AsSecureString
$env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new('', $secureKey).Password
python -m src.asx_maturity_screen --tickers IDX ORA KPG CMW ALQ WGN --force --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output "outputs\asx_maturity_screen.xlsx"
Remove-Item Env:OPENAI_API_KEY
```

Review agreement/tranche separation, amount allocation, exact versus inferred
dates, committed capacity, comparatives, reconciliations, evidence, and flags
before expanding beyond this cohort.

The default first-attempt model is `gpt-5.6-luna`; corrective retries use
`gpt-5.6-terra` and are limited to two after the initial attempt. For a full run
with a usage-based safety estimate:

```powershell
python -m src.asx_maturity_screen --force --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output "outputs\asx_maturity_screen.xlsx" --model gpt-5.6-luna --retry-model gpt-5.6-terra --max-api-cost-usd 20
```

The cost cap is an estimate from returned token usage, not a guarantee of the
final invoice. The workbook is created before the first API request and
atomically checkpointed after every issuer status.
