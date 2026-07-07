"""ke_import.py — 把 ke_backup.py 產生的 JSON 上傳回對應的 KE collection。

- 預設讀 ke/backup/ 下「日期最新」的資料夾，也可用 --dir 指定。
- 檔名（去掉 .json）就是 collection 名稱，必須在 ke/fields 白名單內，否則略過。
- 以 _id 做 ReplaceOne upsert：_id 已存在 → 整筆覆蓋成備份內容；不存在 → 新增。
  重複執行是冪等的，也可用來把備份灌進空的 collection。
- 預設為 DRY-RUN：只預覽每個檔案會新增/覆蓋幾筆，不寫入資料庫。
  實際上傳需加 --execute，並再輸入 y 確認（-y 略過確認）。

用法（從專案根目錄執行）：
  python ke/ke_import.py                          # DRY-RUN 預覽最新備份
  python ke/ke_import.py --dir ke/backup/20260707 # 指定備份資料夾
  python ke/ke_import.py --collections ADHD_F     # 只處理指定 collection
  python ke/ke_import.py --execute                # 預覽後確認，實際上傳
  python ke/ke_import.py --execute -y             # 略過確認直接上傳
"""

import argparse
import glob
import json
import os
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from bson import json_util
from dotenv import load_dotenv
from pymongo import ReplaceOne

from src.importer import get_db

try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

load_dotenv()

FIELDS_GLOB = os.path.join(_SCRIPT_DIR, "fields", "*_ke_fields_grouped.json")
BACKUP_ROOT = os.path.join(_SCRIPT_DIR, "backup")
CHUNK = 1000  # bulk_write / $in 查詢的分批大小


def load_allowed_collections():
    """從最新的 fields JSON 取得 KE 子 collection 白名單（大寫）。"""
    candidates = sorted(glob.glob(FIELDS_GLOB))
    if not candidates:
        sys.exit(f"找不到 fields JSON：{FIELDS_GLOB}")
    path = candidates[-1]
    with open(path, encoding="utf-8") as f:
        grouped = json.load(f)
    print(f"collection 白名單來源：{os.path.basename(path)}（{len(grouped)} 個）")
    return {name.upper() for name in grouped}


def latest_backup_dir():
    dirs = sorted(
        d for d in glob.glob(os.path.join(BACKUP_ROOT, "*"))
        if os.path.isdir(d)
    )
    if not dirs:
        sys.exit(f"ke/backup/ 下沒有任何備份資料夾，請先執行 ke_backup.py")
    return dirs[-1]


def count_existing_ids(db, col_name, ids):
    """回傳 ids 中已存在於 DB 的數量（分批 $in 查詢，唯讀）。"""
    n = 0
    for i in range(0, len(ids), CHUNK):
        n += db[col_name].count_documents({"_id": {"$in": ids[i:i + CHUNK]}})
    return n


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="KE 匯入：把備份 JSON 以 _id upsert 回對應 collection（預設 DRY-RUN）")
    parser.add_argument("--dir", metavar="DIR",
                        help="備份資料夾（預設取 ke/backup/ 下日期最新的一個）")
    parser.add_argument("--collections", nargs="+", metavar="NAME",
                        help="只處理指定的 collection（預設處理資料夾內全部 JSON）")
    parser.add_argument("--execute", action="store_true",
                        help="實際執行上傳（不加此旗標一律只預覽）")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="搭配 --execute 時略過確認直接上傳")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    backup_dir = args.dir if args.dir else latest_backup_dir()
    if not os.path.isdir(backup_dir):
        sys.exit(f"Error: 找不到備份資料夾 {backup_dir}")

    allowed = load_allowed_collections()
    only = {c.upper() for c in args.collections} if args.collections else None

    json_files = sorted(glob.glob(os.path.join(backup_dir, "*.json")))
    if not json_files:
        sys.exit(f"{backup_dir} 內沒有任何 JSON 檔")
    print(f"備份資料夾：{backup_dir}（{len(json_files)} 個 JSON）\n")

    db = get_db()

    # 預覽（唯讀）：每個檔案會新增/覆蓋幾筆
    print("===== 預覽 =====")
    plan = []  # (collection, docs)
    total_new = total_replace = 0
    for path in json_files:
        col = os.path.splitext(os.path.basename(path))[0].upper()
        if only is not None and col not in only:
            continue
        if col not in allowed:
            print(f"  ⚠️  略過 {os.path.basename(path)}：'{col}' 不在 ke_fields 白名單")
            continue

        with open(path, encoding="utf-8") as f:
            docs = json_util.loads(f.read())
        if not isinstance(docs, list):
            print(f"  ⚠️  略過 {os.path.basename(path)}：內容不是 JSON 陣列")
            continue
        if not docs:
            print(f"  [{col}] 檔內 0 筆，略過")
            continue
        if any("_id" not in d for d in docs):
            print(f"  ⚠️  略過 {os.path.basename(path)}：有文件缺 _id，無法以 _id upsert")
            continue

        ids = [d["_id"] for d in docs]
        n_exist = count_existing_ids(db, col, ids)
        n_new = len(docs) - n_exist
        db_count = db[col].count_documents({})
        total_new += n_new
        total_replace += n_exist
        print(f"  [{col}] 檔內 {len(docs)} 筆 → 新增 {n_new}、覆蓋 {n_exist}"
              f"（DB 現有 {db_count} 筆）")
        plan.append((col, docs))

    print(f"\n合計：新增 {total_new} 筆、覆蓋 {total_replace} 筆，"
          f"共 {len(plan)} 個 collection")

    if not args.execute:
        print("DRY-RUN：未寫入任何資料（實際上傳請加 --execute）")
        return
    if not plan:
        print("沒有可上傳的資料，結束。")
        return

    if not args.yes:
        confirm = input(
            f"\n⚠️  確定上傳（新增 {total_new}、覆蓋 {total_replace}）？(y/n)："
        ).strip().lower()
        if confirm != "y":
            print("已取消，未寫入任何資料。")
            return

    print("\n===== 上傳 =====")
    for col, docs in plan:
        upserted = modified = 0
        for i in range(0, len(docs), CHUNK):
            ops = [
                ReplaceOne({"_id": d["_id"]}, d, upsert=True)
                for d in docs[i:i + CHUNK]
            ]
            result = db[col].bulk_write(ops, ordered=False)
            upserted += result.upserted_count
            modified += result.modified_count
        print(f"[{col}] 新增 {upserted} 筆、覆蓋 {modified} 筆")

    print("\ndone")


if __name__ == "__main__":
    main()
