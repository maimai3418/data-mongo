import os
import sys

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from dotenv import load_dotenv
import pandas as pd
from src.importer import get_db
from src.utils.select_collections import select_collections


load_dotenv()

FILEPATH = "delete.xlsx"

def main():
    # 選擇要上傳的量表
    selected = select_collections()
    if selected is None:
        return
    print("reading file...")

    df = pd.read_excel(FILEPATH, dtype=str)
    df.columns = df.columns.str.strip()
    if "record_date" in df.columns:
        df["record_date"] = pd.to_datetime(df["record_date"]).dt.strftime("%Y-%m-%d")
    df = df.where(pd.notna(df), None)
    df = df.dropna(subset=["famid", "record_date", "role"])
    print(f"total rows: {len(df)}")

    db = get_db()
    for col_name in selected:
        filters = []
        for _, row in df.iterrows():
            role_raw = str(row["role"]).strip()
            # "1.0" -> "1"
            try:
                role_val = str(int(float(role_raw)))
            except (ValueError, TypeError):
                role_val = role_raw
            # famid 清理：去頭尾空白、清除 Excel 浮點數殘留（"12345.0" → "12345"）
            famid_val = str(row["famid"]).strip()
            if famid_val.endswith(".0"):
                famid_val = famid_val[:-2]
            f = {
                "famid": famid_val,
                "record_date": str(row["record_date"]).strip(),
                "role": role_val,
            }
            filters.append(f)

        if filters:
            deleted = db[col_name].delete_many({"$or": filters})
            print(f"[{col_name}] deleted {deleted.deleted_count}/{len(filters)}")

    print("done")


if __name__ == "__main__":
    main()