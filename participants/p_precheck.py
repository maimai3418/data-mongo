# p_precheck.py
# 位置：participants/p_precheck.py
# 功能：偵測同一 famid 跨來源 birth_date / sex / group 差異，
#       衝突者掛 diff_dob / diff_sex / diff_group（含對應 *_conflict 旗標）。
# base 規則（每個欄位獨立判定）：
#   DB（Participants）已有值 → 以 DB 值為 base（標記 "base"）；
#   DB 沒有值 → 以第一筆有值的檔案列為 base，並保留真實來源路徑。
# birth_date 統一正規化為 YYYY-MM-DD 後比對；不刪除、不靜默正規化原始欄位。
# 注意：export / write 模式都需要 DB 連線（要先撈既有欄位值當 base）。
#
# 用法（從專案根目錄執行）：
#   python participants/p_precheck.py <檔案或資料夾>
#   python participants/p_precheck.py data_dir/
#   python participants/p_precheck.py a.xlsx b.xlsx
#   python participants/p_precheck.py data_dir/ -o report.xlsx

import os
import sys
import argparse
import pandas as pd
from pathlib import Path
from datetime import date

# 讓 src/ 可被匯入（與 cpt/cpt_precheck.py 相同做法）
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# MongoDB 連線設定與 main.py 相同：讀取專案根目錄 .env 的 MONGO_URI / MONGO_DB
from dotenv import load_dotenv
load_dotenv(os.path.join(_PROJECT_ROOT, ".env"))

# 讓中文/emoji 在 Windows 主控台（cp950）也能正常輸出
try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

# ============ CONFIG ============
FAMID_COL       = "famid"
DOB_COL         = "birth_date"
SEX_COL         = "sex"
GROUP_COL       = "group"
# 檢查欄位 -> (diff 條目內的值 key, diff 欄位, conflict 欄位)
CHECK_FIELDS    = {
    DOB_COL:   ("dob",   "diff_dob",   "dob_conflict"),
    SEX_COL:   ("sex",   "diff_sex",   "sex_conflict"),
    GROUP_COL: ("group", "diff_group", "group_conflict"),
}
SOURCE_PATH_COL = "__source_path"   # 讀檔時填入的來源路徑
TODAY           = date.today().strftime("%Y%m%d")
OUTPUT_FILE     = f"{TODAY}_p_precheck_conflict.xlsx"

# --- MongoDB (mode=write 時使用；連線同 main.py，經 src.importer.get_db) ---
COLLECTION      = "Participants"
# ================================


def _normalize_dob(val) -> str | None:
    """統一為 YYYY-MM-DD；無法解析回傳 None（不丟棄，交由呼叫端決定）。"""
    if val is None or str(val).strip() == "":
        return None
    parsed = pd.to_datetime(str(val).strip(), format="mixed", errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def _normalize_txt(val) -> str | None:
    """sex / group 等文字欄位：去空白、數字尾碼 .0 移除；空值回傳 None。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return None
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def _normalize_for(field: str, val):
    """依欄位選擇正規化方式：birth_date 走日期，其餘走文字。"""
    return _normalize_dob(val) if field == DOB_COL else _normalize_txt(val)


# ── 檔案蒐集與讀取（與 cpt/cpt_precheck.py 相同做法）────────────────────
def collect_files(paths, output_path):
    files = []
    out_ap = os.path.abspath(output_path)
    for p in paths:
        if os.path.isdir(p):
            for f in sorted(os.listdir(p)):
                if f.endswith((".xlsx", ".xls")) and not f.startswith("~$"):
                    files.append(os.path.join(p, f))
        elif os.path.isfile(p):
            files.append(p)
        else:
            print(f"⚠️  略過：找不到 {p}")
    seen, unique = set(), []
    for f in files:
        ap = os.path.abspath(f)
        if ap in seen or ap == out_ap:
            continue
        seen.add(ap)
        unique.append(f)
    return unique


def read_records(files) -> pd.DataFrame:
    """讀取所有檔案的所有工作表，合併為 DataFrame（含 __source_path）。

    欄位名稱統一轉小寫（famid / famID 視為相同）；缺 famid 欄位的工作表略過。
    """
    frames = []
    for filepath in files:
        filename = os.path.basename(filepath)
        try:
            all_sheets = pd.read_excel(filepath, sheet_name=None, dtype=str)
        except Exception as e:
            print(f"⚠️  {filename}: 無法讀取檔案: {e}")
            continue
        for sheet_name, df in all_sheets.items():
            df = df.copy()
            df.columns = df.columns.str.strip().str.lower()
            df = df.loc[:, df.columns.notna()]
            df = df.dropna(how="all")
            if FAMID_COL not in df.columns:
                print(f"⚠️  略過 {filename}｜{sheet_name}：找不到 {FAMID_COL} 欄位")
                continue
            for col in CHECK_FIELDS:
                if col not in df.columns:
                    df[col] = None
            df[SOURCE_PATH_COL] = str(filepath)
            frames.append(df)
    if not frames:
        return pd.DataFrame(columns=[FAMID_COL, *CHECK_FIELDS, SOURCE_PATH_COL])
    return pd.concat(frames, ignore_index=True)


def fetch_db_base() -> dict:
    """從 Participants 撈 famid -> {欄位: 既有值}（僅取有值者，已正規化）。"""
    from src.importer import get_db

    coll = get_db()[COLLECTION]
    proj = {FAMID_COL: 1, "_id": 0}
    proj.update({f: 1 for f in CHECK_FIELDS})
    base = {}
    for doc in coll.find({}, proj):
        fid = doc.get(FAMID_COL)
        if fid is None:
            continue
        fid = str(fid).strip()
        if fid.endswith(".0"):
            fid = fid[:-2]
        vals = {}
        for f in CHECK_FIELDS:
            v = _normalize_for(f, doc.get(f))
            if v is not None:
                vals[f] = v
        if vals:
            base[fid] = vals
    return base


def build_diff_map(records: pd.DataFrame,
                   db_base: dict | None = None) -> pd.DataFrame:
    """
    records 須含：famid, birth_date, sex, group, __source_path
    db_base：famid -> {欄位: DB 既有值（正規化後）}。
    base 規則（每個欄位獨立判定）：
      DB 有值 → DB 值為 base（path/file_name 標記 "base"），
                所有有值的檔案列全部列入比較；
      DB 沒有 → 第一筆有值的檔案列為 base，保留真實來源路徑。
    回傳加上 diff_dob / diff_sex / diff_group（list of dict）
    與 dob_conflict / sex_conflict / group_conflict（bool）的 DataFrame。
    """
    db_base = db_base or {}
    df = records.copy()
    df[FAMID_COL] = df[FAMID_COL].astype(str).str.replace(r"\.0$", "", regex=True)
    for _key, diff_col, conflict_col in CHECK_FIELDS.values():
        df[diff_col] = None
        df[conflict_col] = False

    for famid, grp in df.groupby(FAMID_COL):
        db_vals = db_base.get(famid, {})
        for field, (key, diff_col, conflict_col) in CHECK_FIELDS.items():
            norm = grp[field].apply(lambda v: _normalize_for(field, v))
            valued = grp[norm.notna()]
            if valued.empty:
                continue

            anchor_idx = valued.index[0]   # diff / conflict 掛在此列
            db_val = db_vals.get(field)

            if db_val is not None:
                # DB 已有值 → 以 DB 值為 base，所有檔案列全部列入比較
                entries = [{key: db_val, "path": "base", "file_name": "base"}]
                file_idxs = list(valued.index)
            else:
                # DB 沒有值 → 第一筆有值的檔案列為 base，保留真實來源路徑
                base_src = str(df.at[anchor_idx, SOURCE_PATH_COL])
                entries = [{
                    key: _normalize_for(field, df.at[anchor_idx, field]),
                    "path": base_src,
                    "file_name": Path(base_src).name,
                }]
                file_idxs = [i for i in valued.index if i != anchor_idx]

            for idx in file_idxs:
                val = _normalize_for(field, df.at[idx, field])
                src = str(df.at[idx, SOURCE_PATH_COL])
                entries.append({
                    key: val,
                    "path": src,
                    "file_name": Path(src).name,
                })

            distinct = {e[key] for e in entries}
            df.at[anchor_idx, diff_col] = entries
            df.at[anchor_idx, conflict_col] = len(distinct) > 1

    return df


def export_conflicts(df: pd.DataFrame, output_path: str = OUTPUT_FILE) -> None:
    """僅匯出實際有衝突（任一 *_conflict=True）的記錄供人工檢視。"""
    conflict_cols = [c for _k, _d, c in CHECK_FIELDS.values()]
    diff_cols = [d for _k, d, _c in CHECK_FIELDS.values()]
    flagged = df[df[conflict_cols].any(axis=1)].copy()
    if flagged.empty:
        print("No conflicts found.")
        return
    for col in diff_cols:
        flagged[col] = flagged[col].apply(
            lambda v: str(v) if v is not None else None)
    flagged.to_excel(output_path, index=False)
    print(f"Exported {len(flagged)} conflict record(s) -> {output_path}")


def _merge_entries(old_entries, new_entries, key):
    """合併 DB 既有 diff 與本次 entries（跨次執行累積，不覆蓋）。

    base（path=="base"）以本次為準；其餘以（值, path）去重：
    同一來源重跑不重複累加，新來源接在既有條目之後。
    """
    new_base = next((e for e in new_entries if e.get("path") == "base"), None)
    merged = [new_base] if new_base else []
    seen = set()
    for e in list(old_entries) + list(new_entries):
        if e.get("path") == "base":
            continue
        sig = (e.get(key), e.get("path"))
        if sig in seen:
            continue
        seen.add(sig)
        merged.append(e)
    return merged


def write_to_mongo(df: pd.DataFrame) -> None:
    """把 diff_* / *_conflict 併回 Participants（依 famid 更新）。

    與 DB 既有 diff 合併而非覆蓋；*_conflict 依合併後清單重新判定。
    """
    from pymongo import UpdateOne
    from src.importer import get_db

    diff_cols = [d for _k, d, _c in CHECK_FIELDS.values()]
    rows = [row for _, row in df.iterrows()
            if any(row[d] is not None for d in diff_cols)]
    if not rows:
        print("Nothing to write.")
        return

    coll = get_db()[COLLECTION]

    # 撈這批 famid 既有的 diff 供合併
    famids = sorted({row[FAMID_COL] for row in rows})
    proj = {FAMID_COL: 1, "_id": 0}
    proj.update({d: 1 for d in diff_cols})
    existing = {str(doc[FAMID_COL]): doc
                for doc in coll.find({FAMID_COL: {"$in": famids}}, proj)}

    ops = []
    for row in rows:
        fid = row[FAMID_COL]
        old_doc = existing.get(fid, {})
        update = {}
        for key, diff_col, conflict_col in CHECK_FIELDS.values():
            if row[diff_col] is None:
                continue
            merged = _merge_entries(old_doc.get(diff_col) or [],
                                    row[diff_col], key)
            update[diff_col] = merged
            update[conflict_col] = len({e.get(key) for e in merged}) > 1
        if update:
            ops.append(UpdateOne({FAMID_COL: fid}, {"$set": update}))
    res = coll.bulk_write(ops)
    print(f"Matched {res.matched_count}, modified {res.modified_count}.")


def run(records: pd.DataFrame, mode: str,
        output_path: str = OUTPUT_FILE) -> pd.DataFrame:
    db_base = fetch_db_base()
    print(f"DB 既有基準值 famid 數：{len(db_base)}")
    df = build_diff_map(records, db_base)
    if mode == "write":
        write_to_mongo(df)
    else:
        export_conflicts(df, output_path)
    return df


if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Participants 匯入前檢查：偵測同一 famid 跨來源 "
                    "birth_date / sex / group 差異")
    ap.add_argument("paths", nargs="+",
                    help="要檢查的 xlsx 檔案或資料夾")
    ap.add_argument("-o", "--output", default=OUTPUT_FILE, metavar="XLSX",
                    help=f"報表輸出路徑（預設 {OUTPUT_FILE}）")
    ap.add_argument("--mode", choices=["export", "write"], default="export",
                    help="export=僅匯出檢查檔案；write=寫回 MongoDB")
    args = ap.parse_args()

    files = collect_files(args.paths, args.output)
    if not files:
        print("Error: 找不到任何 xlsx 檔案")
        sys.exit(1)
    print(f"📄 檔案數：{len(files)}")

    records = read_records(files)
    if records.empty:
        print("Error: 沒有讀到任何資料")
        sys.exit(1)
    print(f"共讀取 {len(records)} 列")

    run(records, args.mode, args.output)