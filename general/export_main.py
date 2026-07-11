from dotenv import load_dotenv
from src.importer import get_db
from src.utils.select_collections import select_collections
from datetime import datetime
import json
import os

load_dotenv()


def main():
    selected = select_collections()
    if selected is None:
        return

    db = get_db()
    folder = os.path.join("exports", datetime.now().strftime("%Y%m%d"))
    os.makedirs(folder, exist_ok=True)

    total_exported = 0
    for col_name in selected:
        docs = list(db[col_name].find({}, {"_id": 0}))
        count = len(docs)

        path = os.path.join(folder, f"{col_name}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(docs, f, ensure_ascii=False, indent=2, default=str)

        print(f"[{col_name}] {count} 筆 → {path}")
        total_exported += count

    print(f"\n完成！共匯出 {total_exported} 筆，儲存於 {folder}/")


if __name__ == "__main__":
    main()
