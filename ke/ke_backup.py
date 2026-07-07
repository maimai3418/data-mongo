"""ke_backup.py — 把所有 KE 子 collection 完整下載到 ke/backup/{today}/，每個 collection 一個 JSON。

- collection 清單來自 ke/fields/*_ke_fields_grouped.json（取檔名最新的一份）。
- 只做 find() 讀取，不更動資料庫；輸出寫在本機 ke/backup/{YYYYMMDD}/{COLLECTION}.json。
- 以 bson.json_util 的 Extended JSON 序列化，_id（ObjectId）與日期等 BSON 型態
  都會無損保存，之後可用 ke_import.py 原樣還原。
- 空 collection 也會輸出空陣列 JSON（保留「當時存在但為空」的狀態）。

用法（從專案根目錄執行）：
  python ke/ke_backup.py
"""

import glob
import json
import os
import sys
from datetime import date

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from bson import json_util
from dotenv import load_dotenv

from src.importer import get_db

try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

load_dotenv()

TODAY = date.today().strftime("%Y%m%d")
FIELDS_GLOB = os.path.join(_SCRIPT_DIR, "fields", "*_ke_fields_grouped.json")
BACKUP_DIR = os.path.join(_SCRIPT_DIR, "backup", TODAY)


def load_collection_names():
    """從最新的 fields JSON 取得 KE 子 collection 名稱（大寫）。"""
    candidates = sorted(glob.glob(FIELDS_GLOB))
    if not candidates:
        sys.exit(f"找不到 fields JSON：{FIELDS_GLOB}")
    path = candidates[-1]
    with open(path, encoding="utf-8") as f:
        grouped = json.load(f)
    print(f"collection 清單來源：{os.path.basename(path)}（{len(grouped)} 個）")
    return [name.upper() for name in grouped]


def main():
    db = get_db()
    collections = load_collection_names()

    existing = set(db.list_collection_names())
    missing = [c for c in collections if c not in existing]
    if missing:
        print(f"DB 中不存在、略過（{len(missing)} 個）：{', '.join(missing)}")

    os.makedirs(BACKUP_DIR, exist_ok=True)
    print(f"備份目錄：{BACKUP_DIR}\n")

    total_docs = 0
    saved = 0
    for col in collections:
        if col not in existing:
            continue
        docs = list(db[col].find({}))
        out_path = os.path.join(BACKUP_DIR, f"{col}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(json_util.dumps(docs, ensure_ascii=False, indent=1))
        total_docs += len(docs)
        saved += 1
        print(f"[{col}] {len(docs)} 筆 → {os.path.basename(out_path)}")

    print(f"\ndone，共備份 {saved} 個 collection、{total_docs} 筆到 {BACKUP_DIR}")


if __name__ == "__main__":
    main()
