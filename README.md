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

Schema version 5 separates the raw structured response from a deterministic
normalized payload and separates the report balance date from the screening date. Use
`--as-of-date YYYY-MM-DD` for a reproducible screen (otherwise the local run date
is used) and `--request-timeout-seconds 240` to override the bounded request
timeout. Results created under earlier schema versions are not resumed. The
workbook now includes a seventh **Review Queue** sheet and separate funded-debt
maturity and committed-capacity-expiry grids. Foreign-currency amounts remain in
detail and are excluded from additive reporting-currency totals unless the report
directly provides an equivalent; the pipeline never retrieves exchange rates.

`targets.csv` currently contains 289 data rows and 289 normalized unique
identifiers. The formerly duplicated SGR, RGN, and ACF names are retained as
aliases on single canonical rows; all identifiers pass the 2–5 character local
filename-prefix format. The earlier downloader's reported 291-source-row input
is not stored in this repository, so its 288-identifier count cannot be
reconstructed beyond noting that it was a different source snapshot; no valid
current target was removed to force the older count.

## Post-merge controlled regression

Start with local matching only; this requires no API key:

```powershell
python -m src.asx_maturity_screen --tickers IDX KMD ANN MVF BSL FPR VEA BGA VCX KPG IFL FLT APE ENN --match-only --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output "outputs\asx_maturity_screen_regression.xlsx" --as-of-date 2026-07-29
```

Inspect `outputs\document_register.csv`. The multi-name extraction gate requires
every requested ticker to be `MATCHED_HIGH`; if a genuinely missing ticker is
listed, rerun the same commands with that ticker removed from the `--tickers`
arguments—do not edit Python files. After all remaining documents match, enter
the key without echoing it and run the controlled extraction:

```powershell
$secureKey = Read-Host "OpenAI API key" -AsSecureString
$env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new('', $secureKey).Password
python -m src.asx_maturity_screen --tickers IDX KMD ANN MVF BSL FPR VEA BGA VCX KPG IFL FLT APE ENN --force --pdf-root "C:\Users\chakennedy\2025 FS" --targets targets.csv --output "outputs\asx_maturity_screen_regression.xlsx" --model gpt-5.6-luna --retry-model gpt-5.6-terra --max-api-cost-usd 5 --as-of-date 2026-07-29 --request-timeout-seconds 240
Remove-Item Env:OPENAI_API_KEY
```

Review the regression workbook before considering any full-universe run.
