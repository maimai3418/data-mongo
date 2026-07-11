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
            f = {
                "famid": str(row["famid"]).strip(),
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