#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ADOS 刪除：讀取 ados_delete.xlsx（含 famid、record_date 欄位），
以 famid + record_date 為 unique key，從指定的 ADOS module collection
（ADOS_M1~M4）刪除資料。

與一般 scale 刪除（delete_main.py）的差異：
  - unique key 只看 famid + record_date（**不看 role**），與 ADOS 匯入邏輯一致。
  - 對應 collection 為 ADOS_M1~M4（來自 ados_config.ADOS_COLLECTION_MAP）。

ados_delete.xlsx 欄位：
  - famid（必填）
  - record_date（必填，會正規化為 YYYY-MM-DD 再比對）
  - module（選填）：若該欄存在且填了值（M1~M4），該列只從指定 module 刪除；
                    留空的列則套用到所有選定的 module。

用法（從專案根目錄執行）：
  python ados/ados_delete.py                      # 讀預設 ados_delete.xlsx，互動選 module
  python ados/ados_delete.py -f other.xlsx        # 自訂輸入檔
  python ados/ados_delete.py --modules M1 M3      # 指定 module（略過互動選擇）
  python ados/ados_delete.py --dry-run            # 只預覽符合條件的筆數，不實際刪除
  python ados/ados_delete.py -y                   # 略過確認直接刪除

需求套件：pandas、openpyxl、pymongo、python-dotenv
"""

import os
import sys
import argparse
import warnings

# 讓本腳本可從 ados/ 子目錄被直接執行：把專案根目錄加入 sys.path，
# 以便 import 根目錄的 src 套件（get_db 等）。
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pandas as pd
from dotenv import load_dotenv

from src.importer import get_db
from ados_config import ADOS_COLLECTION_MAP, ADOS_SHARED_FIELDS

# 讓中文/emoji 在 Windows 主控台（cp950）也能正常輸出
try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

load_dotenv()

DEFAULT_FILE = "ados_delete.xlsx"


# ── 讀檔與正規化 ────────────────────────────────────────────────────────
def clean(val):
    """去前後空白、空字串→None，其餘維持原始字串。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    return s or None


def normalize_date(s):
    """把日期字串統一轉成 YYYY-MM-DD 文字；無法解析回 None。"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # 忽略格式推斷警告
        dt = pd.to_datetime(s, errors="coerce")
    return None if pd.isna(dt) else dt.strftime("%Y-%m-%d")


def resolve_file(path):
    """找輸入檔：先看給定路徑（相對 cwd），找不到再退回腳本同層的 ados/。"""
    if os.path.isfile(path):
        return path
    alt = os.path.join(_SCRIPT_DIR, os.path.basename(path))
    return alt if os.path.isfile(alt) else path


def load_delete_rows(filepath):
    """讀 xlsx，回傳 (rows, bad_dates, has_module)。

    rows：[{"famid", "record_date", "module"(或 None)}]，已正規化。
    bad_dates：日期無法解析而被略過的 (famid, 原始日期) 清單。
    """
    df = pd.read_excel(filepath, dtype=str)
    df.columns = df.columns.str.strip()

    for req in ADOS_SHARED_FIELDS:  # famid, record_date
        if req not in df.columns:
            print(f"Error: {os.path.basename(filepath)} 缺少必要欄位 '{req}'")
            sys.exit(1)

    has_module = "module" in df.columns
    rows, bad_dates = [], []
    for _, r in df.iterrows():
        famid = clean(r.get("famid"))
        rd_raw = clean(r.get("record_date"))
        if not famid or not rd_raw:
            continue  # 空列略過
        rd = normalize_date(rd_raw)
        if rd is None:
            bad_dates.append((famid, rd_raw))
            continue
        module = None
        if has_module:
            m = clean(r.get("module"))
            module = m.upper() if m else None
        rows.append({"famid": famid, "record_date": rd, "module": module})
    return rows, bad_dates, has_module


# ── module 選擇 ─────────────────────────────────────────────────────────
def select_modules():
    """互動選擇要刪除的 ADOS module；回傳 module 清單，取消回 None。"""
    available = list(ADOS_COLLECTION_MAP.keys())
    print("\n可用的 ADOS module：")
    for i, m in enumerate(available, 1):
        print(f"  {i}. {m} → {ADOS_COLLECTION_MAP[m]}")
    print(f"  0. 全部 ({', '.join(available)})")
    print("  q. 取消")
    print()

    while True:
        raw = input("請選擇要刪除的 module（編號逗號分隔，例 1,3；0=全部；q=取消）：").strip()
        if raw.lower() == "q":
            print("已取消。")
            return None
        if raw == "0":
            return available
        try:
            indices = [int(x.strip()) for x in raw.split(",")]
            selected = [available[i - 1] for i in indices if 1 <= i <= len(available)]
            if selected:
                print(f"已選擇：{', '.join(selected)}")
                return selected
        except (ValueError, IndexError):
            pass
        print("輸入有誤，請重新輸入。\n")


def filters_for_module(rows, module):
    """為指定 module 組出 famid + record_date 的過濾條件清單。

    module 欄留空的列套用到所有 module；填了其他 module 的列則排除。
    """
    return [
        {"famid": r["famid"], "record_date": r["record_date"]}
        for r in rows
        if r["module"] is None or r["module"] == module
    ]


# ── 主流程 ──────────────────────────────────────────────────────────────
def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="ADOS 刪除：依 famid + record_date 從 ADOS_M1~M4 刪除資料")
    parser.add_argument("-f", "--file", default=DEFAULT_FILE, metavar="XLSX",
                        help=f"輸入 xlsx（預設 {DEFAULT_FILE}）")
    parser.add_argument("--modules", nargs="+", metavar="M",
                        choices=list(ADOS_COLLECTION_MAP.keys()),
                        help="指定 module（略過互動選擇），如 --modules M1 M3")
    parser.add_argument("--dry-run", action="store_true",
                        help="只預覽符合條件的筆數，不實際刪除")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="略過確認直接刪除")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    filepath = resolve_file(args.file)
    if not os.path.isfile(filepath):
        print(f"Error: 找不到檔案 {args.file}")
        sys.exit(1)

    rows, bad_dates, has_module = load_delete_rows(filepath)
    print(f"讀取 {filepath}：有效 {len(rows)} 列"
          + (f"，欄位含 module（可逐列指定）" if has_module else ""))
    for famid, rd in bad_dates:
        print(f"  ⚠️  略過（日期無法解析）：famid={famid}, record_date={rd}")
    if not rows:
        print("沒有可刪除的資料。")
        return

    # 選擇 module
    modules = args.modules if args.modules else select_modules()
    if not modules:
        return

    db = get_db()

    # 先預覽每個 module 符合的筆數
    print("\n===== 預覽 =====")
    plan = []  # (collection, filters, match_count)
    for module in modules:
        col_name = ADOS_COLLECTION_MAP[module]
        filters = filters_for_module(rows, module)
        if not filters:
            print(f"[{col_name}] 無對應列，略過")
            continue
        match = db[col_name].count_documents({"$or": filters})
        print(f"[{col_name}] 符合條件將刪除 {match} 筆（條件 {len(filters)} 組）")
        plan.append((col_name, filters, match))

    total_match = sum(p[2] for p in plan)
    if args.dry_run:
        print(f"\nDRY-RUN：合計符合 {total_match} 筆（未實際刪除）")
        return
    if total_match == 0:
        print("\n沒有符合條件的資料，未執行刪除。")
        return

    # 確認
    if not args.yes:
        confirm = input(
            f"\n⚠️  確定刪除合計 {total_match} 筆？此動作無法復原 (y/n)："
        ).strip().lower()
        if confirm != "y":
            print("已取消，未刪除任何資料。")
            return

    # 實際刪除
    print("\n===== 刪除 =====")
    total_deleted = 0
    for col_name, filters, _ in plan:
        result = db[col_name].delete_many({"$or": filters})
        total_deleted += result.deleted_count
        print(f"[{col_name}] 已刪除 {result.deleted_count} 筆")

    print(f"\ndone，合計刪除 {total_deleted} 筆")


if __name__ == "__main__":
    main()
