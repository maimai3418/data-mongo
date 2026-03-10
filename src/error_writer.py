import pandas as pd
from datetime import datetime

def write_error_xlsx(error_rows: list):
    if not error_rows:
        print("no validation errors")
        return

    df = pd.DataFrame(error_rows)

    # collection 欄位移到最前面，error 欄位移到第二
    cols = ["collection", "error"] + [c for c in df.columns if c not in ("collection", "error")]
    df = df[cols]

    filename = f"errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(filename, index=False)
    print(f"error file saved: {filename}, rows: {len(error_rows)}")