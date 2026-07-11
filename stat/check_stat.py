from dotenv import load_dotenv
from src.importer import get_db

load_dotenv()


def main():
    db = get_db()
    existing = db.list_collection_names()

    while True:
        raw = input("\n請輸入要查詢的 collection 名稱（q 離開）：").strip()
        if raw.lower() == "q":
            print("See you next time!")
            return

        name = raw.upper()
        if name not in existing:
            print(f"找不到 collection：{name}")
            print(f"目前可用的 collections：{', '.join(sorted(existing))}")
            continue

        col = db[name]
        total = col.count_documents({})
        print(f"\n[{name}] 總 document 數量：{total}")

        pipeline = [
            {"$group": {"_id": "$role", "count": {"$sum": 1}}},
            {"$sort": {"_id": 1}},
        ]
        results = list(col.aggregate(pipeline))

        if results:
            print("依 role 分組：")
            for r in results:
                role = r["_id"] if r["_id"] is not None else "(無 role)"
                print(f"  {role}: {r['count']}")
        else:
            print("該 collection 無資料。")


if __name__ == "__main__":
    main()
