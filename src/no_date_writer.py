import pandas as pd
from datetime import datetime
import os

NO_DATE_FILE = "no_date_records.xlsx"


def write_no_date_xlsx(rows: list):
    if not rows:
        print("no age-conflict records to write")
        return

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for row in rows:
        row["timestamp"] = now

    df = pd.DataFrame(rows)
    cols = ["collection", "reason"] + [c for c in df.columns if c not in ("collection", "reason", "timestamp")] + ["timestamp"]
    df = df[cols]

    sheet_name = datetime.now().strftime("%m%d_%H%M%S")

    if os.path.exists(NO_DATE_FILE):
        with pd.ExcelWriter(NO_DATE_FILE, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(NO_DATE_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"no_date_records file updated: {NO_DATE_FILE}, sheet: {sheet_name}, rows: {len(rows)}")
