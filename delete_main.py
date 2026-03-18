from dotenv import load_dotenv
from src.reader import read_xlsx
from src.importer import get_db
from src.utils.select_collections import select_collections


load_dotenv()

FILEPATH = "import_data.xlsx"

def main():
    # 選擇要上傳的量表
    selected = select_collections()
    if selected is None:
        return
    print("reading file...")

    df = read_xlsx(FILEPATH)
    print(f"total rows: {len(df)}")

    db = get_db()
    for col_name in selected:
        filters = []
        for _, row in df.iterrows():
            f = {
                "famid": str(row["famid"]).strip(),
                "record_date": str(row["record_date"]).strip(),
            }
            filters.append(f)

        if filters:
            deleted = db[col_name].delete_many({"$or": filters})
            print(f"[{col_name}] deleted {deleted.deleted_count}/{len(filters)}")

    print("done")


if __name__ == "__main__":
    main()