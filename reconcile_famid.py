#!/usr/bin/env python3
"""
全域 famid 比對（participants 來源 = 一個可抽換的 JSON）。

tool 端：資料夾內所有 .xlsx / .xls，每檔只取第一個 sheet。
participants 端：一個 JSON 檔，結構自動偵測：
    - 陣列            [ {...}, {...} ]
    - JSONL / NDJSON  每行一個物件
    - 包在 key 裡      {"data": [...]} / {"participants": [...]} / {"results": [...]}

雙向比對：
    tool 有、participants 無   -> orphan_famid_report.csv
    participants 有、tool 都無 -> inactive_participant_report.csv

用法：
    python reconcile_famid.py <tool_folder> [participants.json]
    # 第二參數省略時，預設讀執行目錄下的 participants.json

抽換 participants：直接覆蓋那個 json 檔即可，不用改 script。
"""

import sys
import os
import glob
import json
from datetime import datetime, timezone

import pandas as pd

# ----------------------------- CONFIG -----------------------------
FAMID_COL = "famid"
OUTPUT_DIR = "."
DEFAULT_PARTICIPANTS_JSON = "participants.json"   # 預設約定位置
WRAPPER_KEYS = ("data", "participants", "results", "items", "records")
ZERO_PAD = 0   # >0 時把 famid 補零到固定寬度（如 3 -> "002"）；0 表示不補
# ------------------------------------------------------------------


def norm_famid_iter(values) -> set:
    """統一 famid：字串、去空白、去 .0 尾巴、去 NA/空，選擇性補零。"""
    out = set()
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if s == "" or s.lower() in ("nan", "none", "null"):
            continue
        if s.endswith(".0"):
            s = s[:-2]
        if ZERO_PAD > 0:
            s = s.zfill(ZERO_PAD)
        out.add(s)
    return out


# def load_tool_folder(folder: str) -> dict:
    paths = []
    for ext in ("*.xlsx", "*.xls"):
        paths.extend(glob.glob(os.path.join(folder, ext)))
    paths = sorted(set(paths))
    if not paths:
        sys.exit(f"資料夾內找不到 xlsx/xls：{folder}")

    famid_sources = {}
    for p in paths:
        name = os.path.basename(p)
        try:
            df = pd.read_excel(p, sheet_name=0, dtype=str)  # 只取第一個 sheet
        except Exception as e:
            print(f"  [skip] 讀取失敗 {name}: {e}")
            continue
        if FAMID_COL not in df.columns:
            print(f"  [warn] {name} 沒有欄位 '{FAMID_COL}'，跳過")
            continue
        ids = norm_famid_iter(df[FAMID_COL].tolist())
        for fid in ids:
            famid_sources.setdefault(fid, set()).add(name)
        print(f"  [ok] {name}: {len(ids)} 個 famid")
    return famid_sources


def load_tool_folder(folder: str) -> dict:
    paths = []
    for ext in ("*.xlsx", "*.xls"):
        paths.extend(glob.glob(os.path.join(folder, "**", ext), recursive=True))
    paths = sorted(set(paths))
    if not paths:
        sys.exit(f"資料夾內(含子層)找不到 xlsx/xls：{folder}")

    famid_sources = {}
    for p in paths:
        name = os.path.relpath(p, folder)   # 用相對路徑當來源名,避免不同子層同名檔互相蓋掉
        try:
            df = pd.read_excel(p, sheet_name=0, dtype=str)
        except Exception as e:
            print(f"  [skip] 讀取失敗 {name}: {e}")
            continue
        if FAMID_COL not in df.columns:
            print(f"  [warn] {name} 沒有欄位 '{FAMID_COL}'，跳過")
            continue
        ids = norm_famid_iter(df[FAMID_COL].tolist())
        for fid in ids:
            famid_sources.setdefault(fid, set()).add(name)
        print(f"  [ok] {name}: {len(ids)} 個 famid")
    return famid_sources

def load_participants_json(path: str) -> set:
    if not os.path.exists(path):
        sys.exit(f"找不到 participants JSON：{path}")

    raw = open(path, encoding="utf-8").read().strip()

    records = None
    try:
        obj = json.loads(raw)
        if isinstance(obj, list):
            records = obj
        elif isinstance(obj, dict):
            for k in WRAPPER_KEYS:
                if isinstance(obj.get(k), list):
                    records = obj[k]
                    break
            if records is None:
                records = [obj]
    except json.JSONDecodeError:
        records = []
        for line in raw.splitlines():
            line = line.strip().rstrip(",")
            if not line:
                continue
            records.append(json.loads(line))

    if not records:
        sys.exit("participants JSON 解析後沒有任何紀錄，請檢查結構")

    famids = (r.get(FAMID_COL) if isinstance(r, dict) else None for r in records)
    return norm_famid_iter(famids)


def main():
    if len(sys.argv) < 2:
        sys.exit("用法: python reconcile_famid.py <tool_folder> [participants.json]")
    folder = sys.argv[1]
    part_path = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_PARTICIPANTS_JSON

    print(f"讀取 tool 資料夾: {folder}")
    tool_sources = load_tool_folder(folder)
    tool_set = set(tool_sources)
    print(f"tool 端 famid 聯集: {len(tool_set)} 個\n")

    print(f"讀取 participants JSON: {part_path}")
    part_set = load_participants_json(part_path)
    print(f"participants 端 famid: {len(part_set)} 個\n")

    orphan_ids = sorted(tool_set - part_set)
    inactive_ids = sorted(part_set - tool_set)
    valid_ids = tool_set & part_set
    now = datetime.now(timezone.utc).isoformat()

    orphan_df = pd.DataFrame(
        [{"famid": fid, "sources": ",".join(sorted(tool_sources[fid])),
          "status": "pending", "detected_at": now} for fid in orphan_ids]
    )
    inactive_df = pd.DataFrame(
        [{"famid": fid, "status": "pending", "detected_at": now} for fid in inactive_ids]
    )

    orphan_path = os.path.join(OUTPUT_DIR, "orphan_famid_report.csv")
    inactive_path = os.path.join(OUTPUT_DIR, "inactive_participant_report.csv")
    orphan_df.to_csv(orphan_path, index=False, encoding="utf-8-sig")
    inactive_df.to_csv(inactive_path, index=False, encoding="utf-8-sig")

    print("=" * 50)
    print(f"可匯入主表（交集）: {len(valid_ids)}")
    print(f"orphan（tool 有、participants 無）: {len(orphan_ids)} -> {orphan_path}")
    print(f"inactive（participants 有、tool 無）: {len(inactive_ids)} -> {inactive_path}")


if __name__ == "__main__":
    main()
