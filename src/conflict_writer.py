import pandas as pd
from datetime import datetime
import os

CONFLICT_FILE = "conflicts.xlsx"

KEY_COLS = ["collection", "famid", "record_date", "role", "field", "new_value", "db_value"]


def write_conflict_xlsx(conflict_rows: list):
    """value_conflict（key 相同但欄位值不同）寫入 conflicts.xlsx，
    每列一個欄位層級差異（new_value vs db_value 並列），每次執行新增一個時間戳 sheet。"""
    if not conflict_rows:
        print("no value conflicts")
        return

    df = pd.DataFrame(conflict_rows)
    cols = [c for c in KEY_COLS if c in df.columns] + [c for c in df.columns if c not in KEY_COLS]
    df = df[cols]

    sheet_name = datetime.now().strftime("%m%d_%H%M%S")

    if os.path.exists(CONFLICT_FILE):
        with pd.ExcelWriter(CONFLICT_FILE, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(CONFLICT_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"conflict file updated: {CONFLICT_FILE}, sheet: {sheet_name}, rows: {len(conflict_rows)}")
