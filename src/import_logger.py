import pandas as pd
from datetime import datetime
import os

LOG_FILE = "import_log.xlsx"

def write_import_log(log_entries: list):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for entry in log_entries:
        entry["timestamp"] = now

    df_new = pd.DataFrame(log_entries, columns=["collection", "total", "success", "errors", "skipped", "timestamp"])

    if os.path.exists(LOG_FILE):
        df_existing = pd.read_excel(LOG_FILE)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new

    df_combined.to_excel(LOG_FILE, index=False)
    print(f"import log updated: {LOG_FILE}")