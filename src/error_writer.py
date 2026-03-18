import pandas as pd
from datetime import datetime
import os

ERROR_FILE = "errors.xlsx"

def write_error_xlsx(error_rows: list):
    if not error_rows:
        print("no validation errors")
        return

    df = pd.DataFrame(error_rows)
    cols = ["collection", "error"] + [c for c in df.columns if c not in ("collection", "error")]
    df = df[cols]

    sheet_name = datetime.now().strftime("%m%d_%H%M%S")

    if os.path.exists(ERROR_FILE):
        with pd.ExcelWriter(ERROR_FILE, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(ERROR_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"error file updated: {ERROR_FILE}, sheet: {sheet_name}, rows: {len(error_rows)}")