import pandas as pd
from datetime import datetime
import os

LOG_FILE = "import_log.xlsx"

def write_import_log(log_entries: list):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for entry in log_entries:
        entry["timestamp"] = now

    df_new = pd.DataFrame(log_entries)
    col_order = ["collection", "total", "success", "errors", "skipped"]
    extra = [c for c in df_new.columns if c not in col_order and c != "timestamp"]
    col_order = col_order + extra + ["timestamp"]
    df_new = df_new[[c for c in col_order if c in df_new.columns]]

    if os.path.exists(LOG_FILE):
        df_existing = pd.read_excel(LOG_FILE, sheet_name="log")
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new

    # Build pivot: sum numeric columns grouped by collection
    numeric_cols = [c for c in df_combined.columns if c not in ("collection", "timestamp")]
    df_pivot = df_combined.groupby("collection")[numeric_cols].sum().reset_index()

    with pd.ExcelWriter(LOG_FILE, engine="openpyxl") as writer:
        df_combined.to_excel(writer, sheet_name="log", index=False)
        df_pivot.to_excel(writer, sheet_name="pivot", index=False)

    print(f"import log updated: {LOG_FILE}")