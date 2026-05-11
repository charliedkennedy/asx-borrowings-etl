from pathlib import Path
import pandas as pd


def export_all(source_records, borrow_rows, exception_rows, output_dir: Path):
    output_dir.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(source_records).to_csv(output_dir / "source_register.csv", index=False)
    bdf = pd.DataFrame(borrow_rows)
    bdf.to_csv(output_dir / "borrowings_maturity_profile.csv", index=False)
    with pd.ExcelWriter(output_dir / "borrowings_maturity_profile.xlsx", engine="openpyxl") as writer:
        bdf.to_excel(writer, index=False, sheet_name="borrowings")
    pd.DataFrame(exception_rows).to_csv(output_dir / "exceptions_report.csv", index=False)
