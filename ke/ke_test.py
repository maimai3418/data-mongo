"""ke_test.py — 比對兩個 collection 的 unique key（famid + record_date），不更動資料庫。

等同 mongo shell 的：
  keysA/keysB = Set(`${famid}|${record_date}`)
  onlyA = keysA - keysB, onlyB = keysB - keysA

只做 find() 讀取，結果匯出成 xlsx（summary / A_only / B_only 三個分頁）。
"""

import os
import sys
from datetime import date, datetime

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pandas as pd
from dotenv import load_dotenv

from src.importer import get_db
from src.utils.wait_and_retry import wait_and_retry

try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

load_dotenv()

TODAY = date.today().strftime("%Y%m%d")

# ===== CONFIG =====
COLLECTION_A = "ADHD_F"   
COLLECTION_B = "CASE_INFO"
KEY_FIELDS = ["famid", "record_date"]
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           f"{TODAY}_ke_compare.xlsx")
# ===================


def norm(v):
    """統一把 key 欄位轉成可比對的字串。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def fetch_keys(db, collection_name):
    """回傳 (文件總數, {key字串: (famid, record_date)})。"""
    projection = {f: 1 for f in KEY_FIELDS}
    projection["_id"] = 0
    total = 0
    keys = {}
    for doc in db[collection_name].find({}, projection):
        total += 1
        vals = tuple(norm(doc.get(f)) for f in KEY_FIELDS)
        keys["|".join(vals)] = vals
    return total, keys


def to_rows(keys, only):
    return sorted([(*keys[k], k) for k in only])


def main():
    db = get_db()

    total_a, keys_a = fetch_keys(db, COLLECTION_A)
    total_b, keys_b = fetch_keys(db, COLLECTION_B)

    only_a = set(keys_a) - set(keys_b)
    only_b = set(keys_b) - set(keys_a)
    both = set(keys_a) & set(keys_b)

    print(f"[{COLLECTION_A}] docs: {total_a}, unique keys: {len(keys_a)}")
    print(f"[{COLLECTION_B}] docs: {total_b}, unique keys: {len(keys_b)}")
    print(f"A only: {len(only_a)}, B only: {len(only_b)}, both: {len(both)}")
    for k in sorted(only_a)[:10]:
        print("  A:", k)
    for k in sorted(only_b)[:10]:
        print("  B:", k)

    columns = KEY_FIELDS + ["key"]
    summary_df = pd.DataFrame(
        [
            ("A collection", COLLECTION_A),
            ("B collection", COLLECTION_B),
            ("A docs", total_a),
            ("A unique keys", len(keys_a)),
            ("B docs", total_b),
            ("B unique keys", len(keys_b)),
            ("both", len(both)),
            ("A only", len(only_a)),
            ("B only", len(only_b)),
        ],
        columns=["item", "value"],
    )

    def _save():
        with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="summary", index=False)
            pd.DataFrame(to_rows(keys_a, only_a), columns=columns).to_excel(
                writer, sheet_name="A_only", index=False)
            pd.DataFrame(to_rows(keys_b, only_b), columns=columns).to_excel(
                writer, sheet_name="B_only", index=False)

    wait_and_retry(_save, OUTPUT_PATH)

    print(f"已匯出: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
