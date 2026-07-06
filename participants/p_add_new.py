# p_add_new.py
# 位置：participants/p_add_new.py
# 功能：將尚未存在於 Participants 的 famid 送入 staging；
#       staging 內已存在者比對差異並標記，完全相同則跳過。
#       promote 模式：staging 中已補 group 者搬入 Participants。
#
# 用法（從專案根目錄執行）：
#   python participants/p_add_new.py <檔案或資料夾>
#   python participants/p_add_new.py data_dir/
#   python participants/p_add_new.py a.xlsx b.xlsx -o report.xlsx
#   python participants/p_add_new.py --mode promote [--dry-run]

import os
import sys
import argparse
import pandas as pd
from pathlib import Path
from datetime import date, datetime, timezone
from typing import Any

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
COMPARE_FIELDS  = [DOB_COL, SEX_COL]
DIFF_COL        = "diff_fields"
SOURCE_PATH_COL = "__source_path"

TODAY           = date.today().strftime("%Y%m%d")
OUTPUT_FILE     = f"{TODAY}_p_new_candidates.xlsx"

# --- MongoDB（連線同 main.py，經 src.importer.get_db）---
COLL_MAIN       = "Participants"
COLL_STAGING    = "Participants_staging"
# ================================


def _now_utc():
    return datetime.now(timezone.utc)


def _normalize_dob(val) -> str | None:
    if val is None or str(val).strip() == "":
        return None
    parsed = pd.to_datetime(str(val).strip(), format="mixed", errors="coerce")
    return parsed.strftime("%Y-%m-%d") if not pd.isna(parsed) else None


def _normalize_val(field: str, val: Any) -> str | None:
    if field == DOB_COL:
        return _normalize_dob(val)
    if val is None or str(val).strip() == "":
        return None
    return str(val).strip()


def _source_dict(file_path: str) -> dict:
    return {"path": str(file_path), "file_name": Path(file_path).name}


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
            df[SOURCE_PATH_COL] = str(filepath)
            frames.append(df)
    if not frames:
        return pd.DataFrame(columns=[FAMID_COL, SOURCE_PATH_COL])
    return pd.concat(frames, ignore_index=True)


def find_new_famids(records: pd.DataFrame, existing_famids: set) -> pd.DataFrame:
    df = records.copy()
    df[FAMID_COL] = df[FAMID_COL].astype(str).str.replace(r"\.0$", "", regex=True)
    return df[~df[FAMID_COL].isin(existing_famids)]


def reconcile_with_staging(
    new_df: pd.DataFrame,
    staging_docs: dict[str, dict],
) -> tuple[list[dict], list[dict], int]:
    to_insert = []
    to_update = []
    skip_count = 0
    now = _now_utc()

    for _, row in new_df.iterrows():
        fid = row[FAMID_COL]
        src = _source_dict(row[SOURCE_PATH_COL])
        doc = {FAMID_COL: fid}
        for f in COMPARE_FIELDS:
            doc[f] = _normalize_val(f, row.get(f))

        existing = staging_docs.get(fid)
        if existing is None:
            doc[DIFF_COL] = []
            doc["created_at"] = now
            doc["updated_at"] = now
            to_insert.append(doc)
            continue

        diffs_new = []
        all_same = True
        for f in COMPARE_FIELDS:
            inc = _normalize_val(f, row.get(f))
            ext = _normalize_val(f, existing.get(f))
            if inc is not None and ext is not None and inc != ext:
                all_same = False
                diffs_new.append({
                    "field": f, "incoming": inc, "existing": ext, "source": src,
                })
            elif inc is not None and ext is None:
                all_same = False
                diffs_new.append({
                    "field": f, "incoming": inc, "existing": None, "source": src,
                })

        if all_same:
            skip_count += 1
            continue

        prev = existing.get(DIFF_COL, [])
        to_update.append({
            FAMID_COL: fid,
            DIFF_COL: prev + diffs_new,
            "updated_at": now,
        })

    return to_insert, to_update, skip_count


def export_results(to_insert, to_update, skip_count,
                   output_path: str = OUTPUT_FILE) -> None:
    rows = []
    for d in to_insert:
        rows.append({**d, "_action": "insert"})
    for d in to_update:
        rows.append({**d, "_action": "update_diff"})
    if not rows:
        print(f"No new candidates. ({skip_count} skipped as identical)")
        return
    out = pd.DataFrame(rows)
    out[DIFF_COL] = out[DIFF_COL].apply(str)
    out.to_excel(output_path, index=False)
    print(f"Exported {len(to_insert)} insert + {len(to_update)} update -> {output_path}  "
          f"({skip_count} skipped)")


def write_to_staging(to_insert, to_update, skip_count) -> None:
    from pymongo import UpdateOne
    from src.importer import get_db

    coll = get_db()[COLL_STAGING]
    ops = []
    for doc in to_insert:
        ops.append(UpdateOne(
            {FAMID_COL: doc[FAMID_COL]},
            {"$setOnInsert": doc},
            upsert=True,
        ))
    for doc in to_update:
        ops.append(UpdateOne(
            {FAMID_COL: doc[FAMID_COL]},
            {"$set": {DIFF_COL: doc[DIFF_COL], "updated_at": doc["updated_at"]}},
        ))
    if not ops:
        print(f"Nothing to write. ({skip_count} skipped)")
        return
    res = coll.bulk_write(ops)
    print(f"Inserted {res.upserted_count}, modified {res.modified_count}. "
          f"({skip_count} skipped)")


def promote(dry_run: bool = False) -> None:
    """將 staging 中已有 group 的記錄搬入 Participants，搬完從 staging 刪除。"""
    from src.importer import get_db

    db = get_db()
    staging = db[COLL_STAGING]
    main = db[COLL_MAIN]

    ready = list(staging.find({
        GROUP_COL: {"$exists": True, "$ne": None, "$ne": ""},
    }))

    if not ready:
        print("No staging records have group filled. Nothing to promote.")
        return

    no_group = staging.count_documents({
        "$or": [
            {GROUP_COL: {"$exists": False}},
            {GROUP_COL: None},
            {GROUP_COL: ""},
        ]
    })

    print(f"Ready to promote: {len(ready)}  |  Still missing group: {no_group}")

    if dry_run:
        for doc in ready:
            print(f"  [dry-run] {doc[FAMID_COL]}  group={doc[GROUP_COL]}")
        return

    promoted_ids = []
    for doc in ready:
        oid = doc.pop("_id")
        doc.pop(DIFF_COL, None)
        doc["promoted_at"] = _now_utc()
        main.update_one(
            {FAMID_COL: doc[FAMID_COL]},
            {"$setOnInsert": doc},
            upsert=True,
        )
        promoted_ids.append(oid)

    staging.delete_many({"_id": {"$in": promoted_ids}})
    print(f"Promoted {len(promoted_ids)} record(s) to {COLL_MAIN}. "
          f"Remaining in staging: {no_group}")


def run(records: pd.DataFrame, mode: str,
        output_path: str = OUTPUT_FILE) -> None:
    from src.importer import get_db

    db = get_db()

    existing = {
        str(d[FAMID_COL]) for d in db[COLL_MAIN].find({}, {FAMID_COL: 1, "_id": 0})
    }

    new_df = find_new_famids(records, existing)
    if new_df.empty:
        print("All famids already in Participants.")
        return

    staging_docs = {
        str(d[FAMID_COL]): d
        for d in db[COLL_STAGING].find({}, {"_id": 0})
    }

    to_insert, to_update, skip_count = reconcile_with_staging(new_df, staging_docs)

    print(f"New famids not in Participants: {len(new_df)}")
    if mode == "write":
        write_to_staging(to_insert, to_update, skip_count)
    else:
        export_results(to_insert, to_update, skip_count, output_path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Participants 新個案檢查：找出尚未存在的 famid 並送入 staging")
    ap.add_argument("paths", nargs="*",
                    help="要檢查的 xlsx 檔案或資料夾（promote 模式不需要）")
    ap.add_argument("-o", "--output", default=OUTPUT_FILE, metavar="XLSX",
                    help=f"報表輸出路徑（預設 {OUTPUT_FILE}）")
    ap.add_argument("--mode", choices=["export", "write", "promote"], default="export",
                    help="export=匯出檢查檔；write=寫入 staging；promote=staging→Participants")
    ap.add_argument("--dry-run", action="store_true",
                    help="promote 時僅列出符合條件的記錄，不實際搬移")
    args = ap.parse_args()

    if args.mode == "promote":
        promote(dry_run=args.dry_run)
    else:
        if not args.paths:
            ap.error("export/write 模式需要至少一個 xlsx 檔案或資料夾")

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