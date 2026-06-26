#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 CPT 匯出的原始 txt（每筆 8 行的區塊）解析成 xlsx。

用法：
  python cpt_convert.py cpt_raw.txt                 # 預設輸出 cpt_parsed.xlsx
  python cpt_convert.py cpt_raw.txt -o out.xlsx     # 指定輸出檔名
  python cpt_convert.py                             # 讀預設檔 cpt_raw.txt

需求套件：pandas、openpyxl
"""

import os
import sys
import argparse
import warnings

import pandas as pd

warnings.filterwarnings("ignore")

# 讓中文/emoji 在 Windows 主控台（cp950）也能正常輸出
try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

# ── CONFIG ──────────────────────────────────────────────────────────────
DEFAULT_INPUT  = "cpt_raw.txt"
DEFAULT_OUTPUT = "cpt_parsed.xlsx"
LINES_PER_BLOCK = 8
# ────────────────────────────────────────────────────────────────────────

# Line 1 (session), 3–7 的欄位名固定
FIXED_NAMES = {
    1: ["session_dur","s_col2","s_col3","s_col4","location"],
    3: [f"raw_{i+1}" for i in range(15)],
    4: [f"scoreA_{i+1}" for i in range(12)],
    5: [f"scoreB_{i+1}" for i in range(12)],
    6: [f"empty_{i+1}" for i in range(12)],
    7: [f"sum_{i+1}" for i in range(6)],
}

def parse_line(text):
    text = text.strip().strip(",")
    parts = text.split(",")
    while parts and parts[-1].strip() == "":
        parts.pop()
    return [p.strip() for p in parts]

def parse_block(block):
    row = {}
    for line_idx, line in enumerate(block):
        values = parse_line(line)

        # ── Line 0: 自動偵測有無 name 欄 ──
        if line_idx == 0:
            if len(values) >= 9:
                names = ["name","famid","famid_dup","sex","dob","age","code2","test_date","test_time"]
            else:
                names = ["famid","code1","sex","dob","age","code2","test_date","test_time"]

        # ── Line 2: 可能是 '.'、空、或用藥資訊 ──
        elif line_idx == 2:
            meaningful = [v for v in values if v and v != "."]
            row["medication"] = meaningful[0] if meaningful else None
            continue

        # ── Lines 1, 3–7: 用固定欄位名 ──
        else:
            names = FIXED_NAMES.get(line_idx, [f"L{line_idx}_{j+1}" for j in range(len(values))])

        while len(names) < len(values):
            names.append(f"L{line_idx}_{len(names)+1}")
        for name, val in zip(names, values):
            row[name] = val

    return row

def parse_file(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        lines = [l for l in f if l.strip()]
    records = []
    for i in range(0, len(lines), LINES_PER_BLOCK):
        block = lines[i : i + LINES_PER_BLOCK]
        if len(block) < LINES_PER_BLOCK:
            break
        records.append(parse_block(block))
    return records

def build_dataframe(records):
    df = pd.DataFrame(records)

    # famid 字串化 + 去 .0
    for col in ["famid", "famid_dup"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"\.0$", "", regex=True)

    # name 欄：沒有 name 的紀錄填 NaN（混合檔案時會出現）
    # 數值轉換（跳過文字欄）
    skip = {"famid","famid_dup","name","sex","dob","code2","location","medication"}
    for col in df.columns:
        if col in skip:
            continue
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().any():
            df[col] = converted

    # 丟棄全空的 empty_ 欄
    empty_cols = [c for c in df.columns if c.startswith("empty_") and df[c].isna().all()]
    df.drop(columns=empty_cols, inplace=True)

    # 整理欄位順序：把 medication 放在 location 後面
    desired_front = ["name","famid","famid_dup","code1","sex","dob","age","code2",
                     "test_date","test_time","session_dur","s_col2","s_col3","s_col4",
                     "location","medication"]
    front = [c for c in desired_front if c in df.columns]
    rest  = [c for c in df.columns if c not in front]
    return df[front + rest]


def main():
    parser = argparse.ArgumentParser(
        description="把 CPT 匯出的原始 txt 解析成 xlsx")
    parser.add_argument("file", nargs="?", default=DEFAULT_INPUT,
                        help=f"要解析的 CPT txt 檔案路徑（預設 {DEFAULT_INPUT}）")
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help=f"輸出檔名（預設 {DEFAULT_OUTPUT}）")
    args = parser.parse_args()

    path = os.path.abspath(args.file)
    if not os.path.isfile(path):
        print(f"Error: 找不到檔案 {path}")
        sys.exit(1)

    records = parse_file(path)
    df = build_dataframe(records)

    df.to_excel(args.output, index=False, sheet_name="CPT")
    print(f"✅ Done: {len(df)} records → {os.path.abspath(args.output)}")
    print(f"   Columns ({len(df.columns)}): {list(df.columns)}")


if __name__ == "__main__":
    main()