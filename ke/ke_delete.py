"""ke_delete.py — 從指定 KE collection 刪除指定 famid + record_date 的紀錄。

- collection 白名單來自 ke/fields/*_ke_fields_grouped.json（取檔名最新的一份），
  不在白名單內的 collection 整列拒絕，避免誤刪其他量表。
- 輸入固定讀 ke/ke_delete.xlsx 的「delete」worksheet，欄位：
    collection（必填，大小寫皆可，對應 DB 的大寫名稱，如 adhd_f / ADHD_F）
    famid（必填）
    record_date（必填，會正規化成 YYYY-MM-DD 再比對）
- 預設為 DRY-RUN：只逐列預覽 DB 中符合的筆數，不做任何刪除。
  實際刪除需加 --execute，並再輸入 y 確認（-y 略過確認）。

用法（從專案根目錄執行）：
  python ke/ke_delete.py                 # DRY-RUN 預覽（不動資料庫）
  python ke/ke_delete.py -f other.xlsx   # 自訂輸入檔（仍讀 delete 分頁）
  python ke/ke_delete.py --execute       # 預覽後確認，實際刪除
  python ke/ke_delete.py --execute -y    # 略過確認直接刪除
"""

import argparse
import glob
import json
import os
import sys
import warnings

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
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

DEFAULT_FILE = os.path.join(_SCRIPT_DIR, "ke_delete.xlsx")
SHEET_NAME = "delete"
FIELDS_GLOB = os.path.join(_SCRIPT_DIR, "fields", "*_ke_fields_grouped.json")


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


def clean(val):
    """去前後空白、空字串→None，其餘維持原始字串。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    return s or None


def normalize_date(s):
    """把日期字串統一轉成 YYYY-MM-DD 文字；無法解析回 None。"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        dt = pd.to_datetime(s, errors="coerce")
    return None if pd.isna(dt) else dt.strftime("%Y-%m-%d")


def load_delete_rows(filepath, allowed):
    """讀 xlsx 的 delete 分頁，回傳 (有效rows, 被拒絕rows)。

    有效 row：{"collection", "famid", "record_date"}（collection 已轉大寫）。
    被拒絕 row：(列號, 原因, 原始內容摘要)。
    """
    try:
        df = pd.read_excel(filepath, sheet_name=SHEET_NAME, dtype=str)
    except ValueError:
        sys.exit(f"Error: {os.path.basename(filepath)} 沒有名為 '{SHEET_NAME}' 的 worksheet")
    df.columns = df.columns.str.strip()

    for req in ("collection", "famid", "record_date"):
        if req not in df.columns:
            sys.exit(f"Error: '{SHEET_NAME}' 分頁缺少必要欄位 '{req}'")

    rows, rejected = [], []
    for i, r in df.iterrows():
        excel_row = i + 2  # 含標題列的 Excel 實際列號
        col = clean(r.get("collection"))
        famid = clean(r.get("famid"))
        rd_raw = clean(r.get("record_date"))

        if not col and not famid and not rd_raw:
            continue  # 整列空白略過
        if not col or not famid or not rd_raw:
            rejected.append((excel_row, "collection/famid/record_date 有缺",
                             f"{col}, {famid}, {rd_raw}"))
            continue

        col = col.upper()
        if col not in allowed:
            rejected.append((excel_row, "collection 不在 ke_fields 白名單",
                             f"{col}, {famid}, {rd_raw}"))
            continue

        rd = normalize_date(rd_raw)
        if rd is None:
            rejected.append((excel_row, "record_date 無法解析",
                             f"{col}, {famid}, {rd_raw}"))
            continue

        rows.append({"collection": col, "famid": famid, "record_date": rd})
    return rows, rejected


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="KE 刪除：依 collection + famid + record_date 刪除紀錄（預設 DRY-RUN）")
    parser.add_argument("-f", "--file", default=DEFAULT_FILE, metavar="XLSX",
                        help=f"輸入 xlsx（預設 {DEFAULT_FILE}，讀 '{SHEET_NAME}' 分頁）")
    parser.add_argument("--execute", action="store_true",
                        help="實際執行刪除（不加此旗標一律只預覽）")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="搭配 --execute 時略過確認直接刪除")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    if not os.path.isfile(args.file):
        sys.exit(f"Error: 找不到檔案 {args.file}")

    allowed = load_allowed_collections()
    rows, rejected = load_delete_rows(args.file, allowed)

    print(f"讀取 {os.path.basename(args.file)}［{SHEET_NAME}］：有效 {len(rows)} 列、拒絕 {len(rejected)} 列")
    for excel_row, reason, summary in rejected:
        print(f"  ⚠️  第 {excel_row} 列拒絕（{reason}）：{summary}")
    if not rows:
        print("沒有可處理的資料，結束。")
        return

    db = get_db()

    # 逐列預覽 DB 中符合的筆數（唯讀）
    print("\n===== 預覽 =====")
    plan = {}  # collection -> [filter, ...]（只收 DB 有符合的列）
    total_match = 0
    for r in rows:
        flt = {"famid": r["famid"], "record_date": r["record_date"]}
        n = db[r["collection"]].count_documents(flt)
        total_match += n
        mark = "✓" if n else "✗ 無符合"
        print(f"  [{r['collection']}] famid={r['famid']} record_date={r['record_date']} → {n} 筆 {mark}")
        if n:
            plan.setdefault(r["collection"], []).append(flt)

    print(f"\n合計符合 {total_match} 筆")

    if not args.execute:
        print("DRY-RUN：未刪除任何資料（實際刪除請加 --execute）")
        return
    if total_match == 0:
        print("沒有符合條件的資料，未執行刪除。")
        return

    if not args.yes:
        confirm = input(
            f"\n⚠️  確定刪除合計 {total_match} 筆？此動作無法復原 (y/n)："
        ).strip().lower()
        if confirm != "y":
            print("已取消，未刪除任何資料。")
            return

    print("\n===== 刪除 =====")
    total_deleted = 0
    for col_name, filters in plan.items():
        result = db[col_name].delete_many({"$or": filters})
        total_deleted += result.deleted_count
        print(f"[{col_name}] 已刪除 {result.deleted_count} 筆")

    print(f"\ndone，合計刪除 {total_deleted} 筆")


if __name__ == "__main__":
    main()
