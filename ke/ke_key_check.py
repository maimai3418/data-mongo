"""ke_key_check.py — 比對所有 KE 子 collection 的 unique key（famid + record_date）差異，不更動資料庫。

collection 清單來自 ke/fields/*_ke_fields_grouped.json（取檔名最新的一份），
JSON 的 top-level key 為小寫，DB 中 collection 名稱為對應的大寫（asd_s → ASD_S）。

只做 find() 讀取，結果匯出成 xlsx：
  - summary：每個 collection 的文件數、unique key 數、缺少的 key 數
  - diff_matrix：沒有出現在「全部非空」collection 的 key，逐一列出各 collection 有無
    （儲存格為該 key 在該 collection 的筆數，>1 表示 collection 內部有重複；
     空 collection 缺所有 key，只在 summary 標註，不列入差異判斷，以免淹沒真正的差異）
"""

import glob
import json
import os
import sys
from collections import Counter
from datetime import date, datetime

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pandas as pd
from dotenv import load_dotenv

from src.importer import get_db

try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass


load_dotenv()

TODAY = date.today().strftime("%Y%m%d")
KE_DIR = os.path.dirname(os.path.abspath(__file__))

# ===== CONFIG =====
FIELDS_GLOB = os.path.join(KE_DIR, "fields", "*_ke_fields_grouped.json")
KEY_FIELDS = ["famid", "record_date"]
OUTPUT_PATH = os.path.join(KE_DIR, f"{TODAY}_ke_key_check.xlsx")
# ===================


def load_collection_names():
    """從最新的 fields JSON 取得 KE 子 collection 名稱（大寫）。"""
    candidates = sorted(glob.glob(FIELDS_GLOB))
    if not candidates:
        sys.exit(f"找不到 fields JSON：{FIELDS_GLOB}")
    path = candidates[-1]
    with open(path, encoding="utf-8") as f:
        grouped = json.load(f)
    print(f"讀取欄位定義：{os.path.basename(path)}（{len(grouped)} 個 collection）")
    return [name.upper() for name in grouped]


def norm(v):
    """統一把 key 欄位轉成可比對的字串。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def fetch_keys(db, collection_name):
    """回傳 Counter{(famid, record_date): 筆數}。"""
    projection = {f: 1 for f in KEY_FIELDS}
    projection["_id"] = 0
    counter = Counter()
    for doc in db[collection_name].find({}, projection):
        counter[tuple(norm(doc.get(f)) for f in KEY_FIELDS)] += 1
    return counter


def main():
    db = get_db()
    collections = load_collection_names()

    existing = set(db.list_collection_names())
    missing_cols = [c for c in collections if c not in existing]
    found_cols = [c for c in collections if c in existing]
    if missing_cols:
        print(f"DB 中不存在的 collection（{len(missing_cols)} 個）：{', '.join(missing_cols)}")
    if not found_cols:
        sys.exit("沒有任何 KE collection 存在於 DB，結束。")

    keys_by_col = {}
    for col in found_cols:
        keys_by_col[col] = fetch_keys(db, col)
        docs = sum(keys_by_col[col].values())
        print(f"[{col}] docs: {docs}, unique keys: {len(keys_by_col[col])}")

    empty_cols = [c for c in found_cols if not keys_by_col[c]]
    nonempty_cols = [c for c in found_cols if keys_by_col[c]]
    if empty_cols:
        print(f"空的 collection（{len(empty_cols)} 個，不列入差異判斷）：{', '.join(empty_cols)}")
    if not nonempty_cols:
        sys.exit("所有 KE collection 都是空的，結束。")

    all_keys = set().union(*(keys_by_col[c] for c in nonempty_cols))
    print(f"所有 collection 的 key 聯集：{len(all_keys)}")

    # summary
    summary_rows = []
    for col in collections:
        if col not in keys_by_col:
            summary_rows.append((col, "N", "", "", ""))
            continue
        counter = keys_by_col[col]
        summary_rows.append((
            col, "Y",
            sum(counter.values()),          # docs
            len(counter),                   # unique keys
            len(all_keys) - len(counter),   # 缺少的 key 數
        ))
    summary_df = pd.DataFrame(
        summary_rows,
        columns=["collection", "in_db", "docs", "unique_keys", "missing_keys"],
    )

    # diff matrix：只列出沒有出現在全部「非空」collection 的 key
    diff_keys = sorted(k for k in all_keys
                       if any(k not in keys_by_col[c] for c in nonempty_cols))
    diff_rows = []
    for key in diff_keys:
        counts = [keys_by_col[c].get(key, "") for c in nonempty_cols]
        present = sum(1 for c in counts if c != "")
        diff_rows.append((*key, present, *counts))
    diff_df = pd.DataFrame(
        diff_rows,
        columns=KEY_FIELDS + ["present_in"] + nonempty_cols,
    )

    print(f"沒有出現在全部非空 collection（{len(nonempty_cols)} 個）的 key：{len(diff_keys)}")

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        diff_df.to_excel(writer, sheet_name="diff_matrix", index=False)

    print(f"已匯出: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
