#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用 260707_CCPT_merged.xlsx 比對回填 20260702_cpt_cleaned_all.xlsx
（imported_20260703 工作表）缺失的 record_date，另存新檔。

比對鍵（famid 不可靠，不參與比對）：
  tier1：field.json 中除 famid/sex/birth_date/record_date 外的 39 個分數欄全等
  tier2：排除 cleaned 管線重算過的 10 個 ah_* 欄後，其餘 29 個穩定欄全等
每筆配對另附 famid 模糊比對與 age 推算兩個檢核欄，供人工過濾。
已有日期的列不覆寫；配對日期與既有日期不合者出 date_mismatch 報表。

用法：
  python cpt_backfill_dates.py                       # 用 cpt/ 下預設檔
  python cpt_backfill_dates.py --cleaned a.xlsx --merged b.xlsx -o out.xlsx

需求套件：pandas、openpyxl
"""

import os
import re
import sys
import json
import argparse
import warnings

import pandas as pd

from cpt_config import CPT_FIELDS_DIR, discover_field_json

warnings.filterwarnings("ignore")

try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

# ── CONFIG ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CLEANED = os.path.join(SCRIPT_DIR, "20260702_cpt_cleaned_all.xlsx")
DEFAULT_MERGED = os.path.join(SCRIPT_DIR, "260707_CCPT_merged.xlsx")
DEFAULT_OUTPUT_NAME = "20260707_cpt_cleaned_all_backfilled.xlsx"
CLEANED_SHEET = "imported_20260703"
MERGED_SHEET = "merged"

# cleaned 管線重算過的 ah 欄（與 raw 檔必不同，tier2 比對時排除）
AH_RECALC = ["ah_rt", "ah_rtsd", "ah_var", "ah_detect", "ah_rpsty",
             "ah_per", "ah_rtbc", "ah_sebc", "ah_rtisi", "ah_seisi"]

# merged 檔的 famid 變體欄，依可信度排序（取代表 famid 用）
MERGED_FAMID_COLS = ["famid_觀察室_260123", "famid_ADOS_251117", "famid_cpt2_觀察室",
                     "famid_cpt1_ados", "famid_cpt3", "famid_cpt3_raw", "famid_cpt3_t2"]
# ────────────────────────────────────────────────────────────────────────


def load_score_cols():
    """field.json 中除 famid/sex/birth_date/record_date 外的分數欄。"""
    path = discover_field_json(CPT_FIELDS_DIR)
    if not path:
        print("Error: 找不到 CPT field.json")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        fields = json.load(f)
    return [c for c in fields if c not in ("famid", "sex", "birth_date", "record_date")]


def build_key(df, cols):
    """分數欄串成比對鍵。注意 astype(str) 會保留 NaN，必須 fillna。"""
    k = None
    for c in cols:
        part = pd.to_numeric(df[c], errors="coerce").astype(float).round(2) \
                 .astype(str).fillna("nan")
        k = part if k is None else k + "|" + part
    return k


def norm_famid(v):
    """famid 模糊比對用正規化：大寫、去所有非英數字元。"""
    if v is None or pd.isna(v):
        return None
    s = re.sub(r"[^0-9A-Z]", "", str(v).upper())
    return s or None


def norm_date_str(v):
    """任意日期表示 → YYYY-MM-DD 字串，無效回 None。"""
    ts = pd.to_datetime(v, errors="coerce")
    return None if pd.isna(ts) else ts.strftime("%Y-%m-%d")


def build_lookup(mg, key_col):
    """merged 鍵 → 代表列 index。同鍵多列且日期不同者（歧義鍵）整組剔除；
    無 record_date 的列不能拿來回填，也排除。"""
    usable = mg[mg["_rdate"].notna()]
    lookup, dropped = {}, 0
    for k, grp in usable.groupby(key_col):
        if grp["_rdate"].nunique() > 1:
            dropped += 1
            continue
        lookup[k] = grp.index[0]
    return lookup, dropped


def famid_check(cleaned_famid, merged_fams):
    f = norm_famid(cleaned_famid)
    if not f or not merged_fams:
        return "無"
    if f in merged_fams:
        return "同"
    if any(f.startswith(x) or x.startswith(f) for x in merged_fams):
        return "模糊同"
    return "不同"


def age_check(merged_row, cleaned_age):
    try:
        days = (pd.Timestamp(merged_row["_rdate"]) - pd.Timestamp(merged_row["birth_date"])).days
        return "OK" if abs(days // 365 - int(cleaned_age)) <= 1 else "不符"
    except (ValueError, TypeError):
        return "無法算"


def resolve_output(output_arg, base_dir):
    if not output_arg:
        return os.path.join(base_dir, DEFAULT_OUTPUT_NAME)
    out = os.path.abspath(output_arg)
    if os.path.isdir(out) or output_arg.endswith(("/", "\\")):
        return os.path.join(out, DEFAULT_OUTPUT_NAME)
    return out


def main():
    parser = argparse.ArgumentParser(
        description="用 merged 檔比對回填 cleaned 檔缺失的 record_date（另存新檔）")
    parser.add_argument("--cleaned", default=DEFAULT_CLEANED,
                        help=f"cleaned xlsx 路徑（預設 {DEFAULT_CLEANED}）")
    parser.add_argument("--merged", default=DEFAULT_MERGED,
                        help=f"merged xlsx 路徑（預設 {DEFAULT_MERGED}）")
    parser.add_argument("-o", "--output", default=None,
                        help="輸出位置：可給檔案路徑或資料夾"
                             f"（預設輸出到 cleaned 檔所在資料夾，檔名 {DEFAULT_OUTPUT_NAME}）")
    args = parser.parse_args()

    for p in (args.cleaned, args.merged):
        if not os.path.isfile(p):
            print(f"Error: 找不到檔案 {os.path.abspath(p)}")
            sys.exit(1)
    out_path = resolve_output(args.output, os.path.dirname(os.path.abspath(args.cleaned)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    score_cols = load_score_cols()
    stable_cols = [c for c in score_cols if c not in AH_RECALC]

    imp = pd.read_excel(args.cleaned, sheet_name=CLEANED_SHEET)
    mg = pd.read_excel(args.merged, sheet_name=MERGED_SHEET)
    orig_cols = list(imp.columns)

    imp["_k39"] = build_key(imp, score_cols)
    imp["_k29"] = build_key(imp, stable_cols)
    mg["_k39"] = build_key(mg, score_cols)
    mg["_k29"] = build_key(mg, stable_cols)
    mg["_rdate"] = mg["record_date"].map(norm_date_str)
    mg["_fams"] = mg[MERGED_FAMID_COLS].apply(
        lambda r: {norm_famid(v) for v in r if norm_famid(v)}, axis=1)
    mg["_famid_disp"] = mg[MERGED_FAMID_COLS].apply(
        lambda r: next((str(v) for v in r if pd.notna(v)), None), axis=1)

    map39, amb39 = build_lookup(mg, "_k39")
    map29, amb29 = build_lookup(mg, "_k29")

    # 逐列比對
    new_cols = ["backfill", "backfill_tier", "backfill_date", "backfill_merged_famid",
                "backfill_famid_check", "backfill_age_check", "date_mismatch"]
    for c in new_cols:
        imp[c] = ""
    mismatch_rows = []
    stats = {"tier1": 0, "tier2": 0, "unmatched": 0, "mismatch": 0}
    fam_stats, age_stats = {}, {}

    for i, r in imp.iterrows():
        if r["_k39"] in map39:
            tier, mi = "39欄全等", map39[r["_k39"]]
        elif r["_k29"] in map29:
            tier, mi = "29欄全等_ah重算", map29[r["_k29"]]
        else:
            if pd.isna(r["record_date"]):
                stats["unmatched"] += 1
            continue
        m = mg.loc[mi]

        if pd.notna(r["record_date"]):
            # 已有日期：不動，只檢查是否吻合
            if norm_date_str(r["record_date"]) != m["_rdate"]:
                imp.at[i, "date_mismatch"] = "Y"
                stats["mismatch"] += 1
                mismatch_rows.append({
                    "excel_row": i + 2,
                    "famid_cleaned": r["famid"],
                    "famid_merged": m["_famid_disp"],
                    "record_date_cleaned": norm_date_str(r["record_date"]),
                    "record_date_merged": m["_rdate"],
                    "age_cleaned": r["age"],
                    "tier": tier,
                    "cleaned_狀態": r["cleaned_狀態"],
                })
            continue

        # 缺日期：回填 + 檢核欄
        fc = famid_check(r["famid"], m["_fams"])
        ac = age_check(m, r["age"])
        imp.at[i, "record_date"] = pd.Timestamp(m["_rdate"])
        imp.at[i, "backfill"] = "Y"
        imp.at[i, "backfill_tier"] = tier
        imp.at[i, "backfill_date"] = m["_rdate"]
        imp.at[i, "backfill_merged_famid"] = m["_famid_disp"]
        imp.at[i, "backfill_famid_check"] = fc
        imp.at[i, "backfill_age_check"] = ac
        key = "tier1" if tier == "39欄全等" else "tier2"
        stats[key] += 1
        fam_stats[(key, fc)] = fam_stats.get((key, fc), 0) + 1
        age_stats[(key, ac)] = age_stats.get((key, ac), 0) + 1

    # 輸出
    out_main = imp[orig_cols + new_cols]
    unmatched = imp[(imp.record_date.isna())]
    unmatched_out = unmatched[["famid", "age", "sex", "record_date_raw", "cleaned_狀態",
                               "rn_omis", "rp_omis", "rn_comis", "rp_comis", "r_rt"]].copy()
    unmatched_out.insert(0, "excel_row", unmatched.index + 2)
    mismatch_df = pd.DataFrame(mismatch_rows)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_main.to_excel(writer, index=False, sheet_name="imported_backfilled")
        mismatch_df.to_excel(writer, index=False, sheet_name="date_mismatch")
        unmatched_out.to_excel(writer, index=False, sheet_name="unmatched")

    n_miss = stats["tier1"] + stats["tier2"] + stats["unmatched"]
    print(f"✅ Done → {out_path}")
    print(f"   缺 record_date 共 {n_miss} 列："
          f"tier1(39欄全等) 回填 {stats['tier1']}、tier2(29欄全等) 回填 {stats['tier2']}、"
          f"未匹配 {stats['unmatched']}")
    print(f"   已有日期但與 merged 不合（date_mismatch）: {stats['mismatch']} 列")
    print(f"   歧義鍵剔除：k39 {amb39} 個、k29 {amb29} 個")
    for key in ("tier1", "tier2"):
        fam = {k[1]: v for k, v in fam_stats.items() if k[0] == key}
        age = {k[1]: v for k, v in age_stats.items() if k[0] == key}
        print(f"   {key} famid 檢核: {fam} | age 檢核: {age}")


if __name__ == "__main__":
    main()
