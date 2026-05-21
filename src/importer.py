from pymongo import MongoClient, UpdateOne
from config import  SHARED_FIELDS
import os

def get_db():
    client = MongoClient(os.getenv("MONGO_URI"))
    return client[os.getenv("MONGO_DB")]

def upsert_many(db, collection_name: str, docs: list):
    if not docs:
        print(f"[{collection_name}] no data, skipped")
        return

    col = db[collection_name]
    operations = [
        UpdateOne(
            {k: doc[k] for k in SHARED_FIELDS if k in doc},
            {"$setOnInsert": doc},
            upsert=True
        )
        for doc in docs
    ]
    result = col.bulk_write(operations)
    skipped = len(docs) - result.upserted_count
    print(f"[{collection_name}] inserted: {result.upserted_count}, skipped (already exists): {skipped}")
    return result.upserted_count, skipped