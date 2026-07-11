from pymongo import MongoClient, UpdateOne
from config import  SHARED_FIELDS
import os

def get_db():
    client = MongoClient(os.getenv("MONGO_URI"))
    return client[os.getenv("MONGO_DB")]

def upsert_many(db, collection_name: str, docs: list):
    """依 composite key（SHARED_FIELDS）比對 DB 現有紀錄，分類處理：
    - new_insert：DB 無對應 key → 插入
    - exact_dup：key 相同且新資料所有欄位值與 DB 相同 → 跳過（計數）
    - value_conflict：key 相同但欄位值不同 → 不覆蓋 DB，回傳欄位層級差異供匯出衝突報告

    Returns: (inserted, exact_dup, conflict_docs, conflict_rows)
    """
    if not docs:
        print(f"[{collection_name}] no data, skipped")
        return

    col = db[collection_name]

    def key_of(doc):
        return tuple(doc.get(k) for k in SHARED_FIELDS)

    # 一次撈出 DB 中 key 相同的既有紀錄
    key_filters = [{k: doc.get(k) for k in SHARED_FIELDS} for doc in docs]
    existing = {}
    for db_doc in col.find({"$or": key_filters}):
        existing[tuple(db_doc.get(k) for k in SHARED_FIELDS)] = db_doc

    to_insert = []
    exact_dup = 0
    conflict_docs = 0
    conflict_rows = []

    for doc in docs:
        db_doc = existing.get(key_of(doc))
        if db_doc is None:
            to_insert.append(doc)
            continue

        # key 相同 → 逐欄比對新資料的非 key 欄位
        diff_fields = [
            f for f in doc
            if f not in SHARED_FIELDS and doc.get(f) != db_doc.get(f)
        ]
        if diff_fields:
            conflict_docs += 1
            for f in diff_fields:
                conflict_rows.append({
                    "collection": collection_name,
                    **{k: doc.get(k) for k in SHARED_FIELDS},
                    "field": f,
                    "new_value": doc.get(f),
                    "db_value": db_doc.get(f),
                })
        else:
            exact_dup += 1

    inserted = 0
    if to_insert:
        # 仍用 $setOnInsert upsert：批次內 key 重複時只插入第一筆
        operations = [
            UpdateOne(
                {k: doc[k] for k in SHARED_FIELDS if k in doc},
                {"$setOnInsert": doc},
                upsert=True
            )
            for doc in to_insert
        ]
        result = col.bulk_write(operations)
        inserted = result.upserted_count

    print(
        f"[{collection_name}] inserted: {inserted}, "
        f"exact_dup (skipped): {exact_dup}, "
        f"value_conflict (skipped, see report): {conflict_docs}"
    )
    return inserted, exact_dup, conflict_docs, conflict_rows
