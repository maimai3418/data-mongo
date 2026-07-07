#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 260707_CCPT_raw_all_test.xlsx 的 5 張工作表合併去重成單一檔案。

famid 不完全可靠，因此以 field.json 中除 famid/sex/birth_date 外的欄位
（record_date + 39 個分數欄）當指紋去重：同指紋視為同一筆施測。
各來源的 famid 並排保留（不選主 famid），同指紋但 famid 不一致的組
另開 famid_conflicts 工作表逐來源展開，供人工核對。

用法：
  python cpt_merge_raw.py                            # 讀 cpt/ 下預設檔
  python cpt_merge_raw.py input.xlsx -o out.xlsx     # 自訂輸入/輸出位置
  python cpt_merge_raw.py -o C:\results              # 輸出到指定資料夾

需求套件：pandas、openpyxl
"""

import os
import sys
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
DEFAULT_INPUT = os.path.join(SCRIPT_DIR, "260707_CCPT_raw_all_test.xlsx")
DEFAULT_OUTPUT_NAME = "260707_CCPT_merged.xlsx"

# 來源工作表；PRIORITY 順序 = 取欄位值時的優先序（metadata 最完整者優先）
PRIORITY = [
    "CPT_觀察室_260123",
    "CPT_ADOS_251117_score",
    "cpt2_觀察室",
    "cpt1_ados",
    "cpt3_斷頭",
]
# 各表 famid 欄在輸出中的並排欄名
FAMID_OUT_COL = {
    "CPT_觀察室_260123":     "famid_觀察室_260123",
    "CPT_ADOS_251117_score": "famid_ADOS_251117",
    "cpt2_觀察室":           "famid_cpt2_觀察室",
    "cpt1_ados":             "famid_cpt1_ados",
    "cpt3_斷頭":             "famid_cpt3",
}
# 判斷 test_type 用：cpt3_斷頭 不算（僅斷頭時另標）
ADOS_SHEETS = {"cpt1_ados", "CPT_ADOS_251117_score"}
OBS_SHEETS  = {"cpt2_觀察室", "CPT_觀察室_260123"}
# ────────────────────────────────────────────────────────────────────────


def load_score_cols():
    """從 field.json 取出指紋欄位：除 famid/sex/birth_date/record_date 外的分數欄。"""
    import json
    path = discover_field_json(CPT_FIELDS_DIR)
    if not path:
        print("Error: 找不到 CPT field.json")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        fields = json.load(f)
    return [c for c in fields if c not in ("famid", "sex", "birth_date", "record_date")]


def norm_date(series):
    """任何日期表示（datetime / MM/DD/YYYY / 0000-00-00…）→ YYYY-MM-DD 字串，無效為 NaN。"""
    return pd.to_datetime(series, errors="coerce").dt.strftime("%Y-%m-%d")


def norm_id(series):
    """famid 類欄位：字串化、strip、去尾端 .0，空值統一為 NaN。"""
    s = series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return s.where(~s.isin(["nan", "NaN", "None", ""]), None)


def norm_sex(series):
    """M/F 或 1/2 → 1/2（M=1、F=2，已用配對資料驗證）。"""
    s = series.astype(str).str.strip().str.upper().str.replace(r"\.0$", "", regex=True)
    return s.map({"M": "1", "F": "2", "1": "1", "2": "2"})


def fingerprint(df, score_cols):
    """record_date + 39 分數欄（round 3）串成指紋。
    注意：astype(str) 會把 NaN 保留成 NaN（非 'nan' 字串），必須 fillna。"""
    fp = norm_date(df["record_date"]).fillna("NA")
    for c in score_cols:
        part = df[c].round(3).astype(str).fillna("nan")
        fp = fp + "|" + part
    return fp


def load_sheets(path, score_cols):
    """讀取並正規化 5 張工作表，回傳 {sheet: DataFrame}（含 _fp/_famid/_sex/_birth/_rdate）。"""
    xl = pd.ExcelFile(path)
    missing = [s for s in PRIORITY if s not in xl.sheet_names]
    if missing:
        print(f"Error: 輸入檔缺少工作表 {missing}，實際有 {xl.sheet_names}")
        sys.exit(1)

    frames = {}
    for sheet in PRIORITY:
        df = xl.parse(sheet)
        for c in score_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
        df["_fp"] = fingerprint(df, score_cols)
        df["_rdate"] = norm_date(df["record_date"])
        df["_famid"] = norm_id(df["famid"])
        df["_sex"] = norm_sex(df["sex"])
        # 兩種 schema 的生日欄不同：cpt* 用 birth_date、CPT_* 用 dob
        birth_src = df["birth_date"] if "birth_date" in df.columns else df["dob"]
        df["_birth"] = norm_date(birth_src)
        if sheet == "cpt3_斷頭":
            df["_famid_raw"] = norm_id(df["famid_raw"])
            df["_famid_t2"] = norm_id(df["famid_t2"])
        frames[sheet] = df
    return frames


def distinct(values):
    return sorted(set(v for v in values if v is not None and pd.notna(v)))


def merge_groups(frames, score_cols):
    """把各表列依指紋分組，每組合併成一列。回傳 (merged_rows, conflict_rows)。"""
    # fp → {sheet: row_dict}（已驗證各表內部零重複，保險起見仍只取第一筆）
    groups = {}
    for sheet in PRIORITY:
        for rec in frames[sheet].to_dict("records"):
            groups.setdefault(rec["_fp"], {}).setdefault(sheet, rec)

    def pick(rows, key):
        """依 PRIORITY 取第一個非空值。"""
        for sheet in PRIORITY:
            rec = rows.get(sheet)
            if rec is not None and key in rec and pd.notna(rec.get(key)):
                return rec[key]
        return None

    merged, conflicts = [], []
    for fp, rows in groups.items():
        sheets = [s for s in PRIORITY if s in rows]
        famids = distinct(rows[s]["_famid"] for s in sheets)
        famid_conflict = len(famids) > 1
        sexes = distinct(rows[s]["_sex"] for s in sheets)
        births = distinct(rows[s]["_birth"] for s in sheets)

        if set(sheets) & ADOS_SHEETS:
            test_type = "ADOS"
        elif set(sheets) & OBS_SHEETS:
            test_type = "觀察室"
        else:
            test_type = "僅斷頭"

        out = {}
        for sheet in PRIORITY:
            rec = rows.get(sheet)
            out[FAMID_OUT_COL[sheet]] = rec["_famid"] if rec else None
        c3 = rows.get("cpt3_斷頭")
        out["famid_cpt3_raw"] = c3["_famid_raw"] if c3 else None
        out["famid_cpt3_t2"] = c3["_famid_t2"] if c3 else None

        out["sex"] = sexes[0] if len(sexes) == 1 else pick(rows, "_sex")
        out["birth_date"] = births[0] if len(births) == 1 else pick(rows, "_birth")
        out["record_date"] = pick(rows, "_rdate")
        out["test_type"] = test_type
        for meta in ["name", "famid_dup", "cpt_age", "test_date", "test_time",
                     "session_dur", "流水號"]:
            out[meta] = pick(rows, meta)
        for c in score_cols:
            out[c] = pick(rows, c)
        out["sources"] = ";".join(sheets)
        out["n_sources"] = len(sheets)
        out["famid_conflict"] = "Y" if famid_conflict else ""
        out["sex_conflict"] = "Y" if len(sexes) > 1 else ""
        out["birth_date_conflict"] = "Y" if len(births) > 1 else ""
        merged.append(out)

        if famid_conflict:
            conflicts.append((fp, rows, sheets))

    return merged, conflicts


def build_merged_df(merged, score_cols):
    df = pd.DataFrame(merged)
    front = (list(FAMID_OUT_COL.values()) + ["famid_cpt3_raw", "famid_cpt3_t2",
             "sex", "birth_date", "record_date", "test_type", "name", "famid_dup",
             "cpt_age", "test_date", "test_time", "session_dur", "流水號"])
    tail = ["sources", "n_sources", "famid_conflict", "sex_conflict", "birth_date_conflict"]
    df = df[front + score_cols + tail]
    df["sex"] = pd.to_numeric(df["sex"], errors="coerce").astype("Int64")
    return df.sort_values(["record_date", "famid_觀察室_260123"],
                          na_position="last", kind="stable").reset_index(drop=True)


def build_conflicts_df(conflicts):
    """衝突組逐來源展開，每來源一列。"""
    conflicts = sorted(conflicts, key=lambda t: (t[1][t[2][0]]["_rdate"] or "", t[0]))
    rows = []
    for gid, (fp, group, sheets) in enumerate(conflicts, start=1):
        for sheet in sheets:
            rec = group[sheet]
            rows.append({
                "group_id": gid,
                "record_date": rec["_rdate"],
                "sheet": sheet,
                "famid": rec["_famid"],
                "famid_raw": rec.get("_famid_raw"),
                "famid_t2": rec.get("_famid_t2"),
                "name": rec.get("name"),
                "sex": rec["_sex"],
                "birth_date": rec["_birth"],
                "rn_omis": rec.get("rn_omis"),
                "rp_omis": rec.get("rp_omis"),
                "r_rt": rec.get("r_rt"),
            })
    df = pd.DataFrame(rows)
    df["sex"] = pd.to_numeric(df["sex"], errors="coerce").astype("Int64")
    return df


def resolve_output(output_arg, input_path):
    if not output_arg:
        return os.path.join(os.path.dirname(input_path), DEFAULT_OUTPUT_NAME)
    out = os.path.abspath(output_arg)
    if os.path.isdir(out) or output_arg.endswith(("/", "\\")):
        return os.path.join(out, DEFAULT_OUTPUT_NAME)
    return out


def main():
    parser = argparse.ArgumentParser(
        description="合併去重 260707_CCPT_raw_all_test.xlsx 的 5 張工作表")
    parser.add_argument("file", nargs="?", default=DEFAULT_INPUT,
                        help=f"輸入 xlsx 路徑（預設 {DEFAULT_INPUT}）")
    parser.add_argument("-o", "--output", default=None,
                        help="輸出位置：可給檔案路徑或資料夾"
                             f"（預設輸出到輸入檔所在資料夾，檔名 {DEFAULT_OUTPUT_NAME}）")
    args = parser.parse_args()

    path = os.path.abspath(args.file)
    if not os.path.isfile(path):
        print(f"Error: 找不到檔案 {path}")
        sys.exit(1)
    out_path = resolve_output(args.output, path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    score_cols = load_score_cols()
    frames = load_sheets(path, score_cols)
    total_rows = sum(len(df) for df in frames.values())

    merged, conflicts = merge_groups(frames, score_cols)
    merged_df = build_merged_df(merged, score_cols)
    conflicts_df = build_conflicts_df(conflicts)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="merged")
        conflicts_df.to_excel(writer, index=False, sheet_name="famid_conflicts")

    print(f"✅ Done: {total_rows} 列（5 表）→ {len(merged_df)} 筆唯一施測 → {out_path}")
    for sheet in PRIORITY:
        n = len(frames[sheet])
        only = sum(1 for m in merged if m["sources"] == sheet)
        print(f"   [{sheet}] {n} 列，其中 {only} 筆僅此表獨有")
    print(f"   famid 衝突組數: {len(conflicts)}（famid_conflicts 工作表共 {len(conflicts_df)} 列）")
    print(f"   sex 衝突: {(merged_df['sex_conflict'] == 'Y').sum()} 筆、"
          f"birth_date 衝突: {(merged_df['birth_date_conflict'] == 'Y').sum()} 筆")


if __name__ == "__main__":
    main()
