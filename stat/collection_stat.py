from collections import Counter
from datetime import datetime
import os
import sys

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pandas as pd
from dotenv import load_dotenv
from src.importer import get_db
from config import COLLECTION_MAP

load_dotenv()

FIVE_DIGIT_PREFIXES = ("3", "5", "6", "8", "9")
SIX_DIGIT_PREFIXES = ("4",)
ONE_PREFIX_LEN_RANGE = (2, 5)  # 1 開頭、長度 2~5 位
EXPORT_DIR = os.path.join("exports", "collection_stat")
EXPORT_STEM = "collection_stat"
EXPORT_EXT = ".xlsx"

BUCKET_ONE = "1 開頭(2-5位)"
LAST_DIGITS = tuple(str(i) for i in range(10))  # "0".."9"


def fetch_famids(col):
    famids = []
    for doc in col.find({}, {"famid": 1, "_id": 0}):
        f = doc.get("famid")
        if f is None:
            continue
        s = str(f).strip()
        if s == "":
            continue
        famids.append(s)
    return famids


def classify(fid: str) -> str | None:
    if not fid.isdigit():
        return None
    lo, hi = ONE_PREFIX_LEN_RANGE
    if fid[0] == "1" and lo <= len(fid) <= hi:
        return BUCKET_ONE
    if len(fid) == 5 and fid[0] in FIVE_DIGIT_PREFIXES:
        return f"{fid[0]} 開頭(5位)"
    if len(fid) == 6 and fid[0] in SIX_DIGIT_PREFIXES:
        return f"{fid[0]} 開頭(6位)"
    return None


def bucket_keys() -> list[str]:
    keys = [BUCKET_ONE]
    keys += [f"{p} 開頭(5位)" for p in FIVE_DIGIT_PREFIXES]
    keys += [f"{p} 開頭(6位)" for p in SIX_DIGIT_PREFIXES]
    return keys


def compute_stats(famids: list[str]) -> dict:
    distinct = set(famids)
    buckets = {k: 0 for k in bucket_keys()}
    last_digit = {d: 0 for d in LAST_DIGITS}
    last_digit_other = 0
    other = 0

    for fid in distinct:
        key = classify(fid)
        if key is None:
            other += 1
        else:
            buckets[key] += 1

        if fid and fid[-1] in last_digit:
            last_digit[fid[-1]] += 1
        else:
            last_digit_other += 1

    appearance = Counter(famids)
    distribution = Counter(appearance.values())

    return {
        "total_records": len(famids),
        "distinct_famid": len(distinct),
        "buckets": buckets,
        "other": other,
        "last_digit": last_digit,
        "last_digit_other": last_digit_other,
        "distribution": distribution,
    }


def print_stats(name: str, stats: dict):
    print(f"\n[{name}] 總紀錄數（含 famid）: {stats['total_records']}")
    print(f"distinct famid 數量: {stats['distinct_famid']}")

    print("依 famid 格式分類（distinct famid）：")
    for k, v in stats["buckets"].items():
        print(f"  {k}: {v}")
    print(f"  其他: {stats['other']}")
    print(f"  目標分類合計: {sum(stats['buckets'].values())}")

    print("依 famid 結尾數字分類（distinct famid）：")
    for d in LAST_DIGITS:
        print(f"  {d} 結尾: {stats['last_digit'][d]}")
    if stats["last_digit_other"]:
        print(f"  其他(非數字結尾): {stats['last_digit_other']}")

    print("famid 重複出現次數分布：")
    for n in sorted(stats["distribution"].keys()):
        print(f"  count={n}: {stats['distribution'][n]} 個 famid")


def export_all(db):
    existing = set(db.list_collection_names())
    bucket_cols = bucket_keys()
    last_cols = [f"{d} 結尾" for d in LAST_DIGITS]

    summary_rows = []
    last_rows = []
    dup_rows = []
    max_count = 0

    targets = list(COLLECTION_MAP.keys())
    print(f"\n掃描 {len(targets)} 個 collection（依 COLLECTION_MAP）...")

    for name in targets:
        if name not in existing:
            print(f"  [{name}] 不存在於 DB，跳過")
            continue

        famids = fetch_famids(db[name])
        stats = compute_stats(famids)

        row = {
            "collection": name,
            "total_records": stats["total_records"],
            "distinct_famid": stats["distinct_famid"],
        }
        row.update(stats["buckets"])
        row["其他"] = stats["other"]
        summary_rows.append(row)

        last_row = {"collection": name, "distinct_famid": stats["distinct_famid"]}
        for d in LAST_DIGITS:
            last_row[f"{d} 結尾"] = stats["last_digit"][d]
        last_row["其他(非數字結尾)"] = stats["last_digit_other"]
        last_rows.append(last_row)

        dup_row = {"collection": name}
        for n, cnt in stats["distribution"].items():
            dup_row[n] = cnt
            if n > max_count:
                max_count = n
        dup_rows.append(dup_row)

        print(f"  [{name}] distinct={stats['distinct_famid']}, total={stats['total_records']}")

    summary_cols = ["collection", "total_records", "distinct_famid"] + bucket_cols + ["其他"]
    df_summary = pd.DataFrame(summary_rows, columns=summary_cols)

    last_full_cols = ["collection", "distinct_famid"] + last_cols + ["其他(非數字結尾)"]
    df_last = pd.DataFrame(last_rows, columns=last_full_cols)

    count_cols = list(range(1, max_count + 1))
    df_dup = pd.DataFrame(dup_rows).reindex(columns=["collection"] + count_cols).fillna(0)
    df_dup.columns = ["collection"] + [f"count={n}" for n in count_cols]
    int_cols = [c for c in df_dup.columns if c != "collection"]
    df_dup[int_cols] = df_dup[int_cols].astype(int)

    os.makedirs(EXPORT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    db_name = os.getenv("MONGO_DB")
    out_path = os.path.join(
        EXPORT_DIR,
        f"{ts}_{EXPORT_STEM}_{db_name}{EXPORT_EXT}",
    )
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="summary", index=False)
        df_last.to_excel(writer, sheet_name="by_last_digit", index=False)
        df_dup.to_excel(writer, sheet_name="duplicates", index=False)

    print(f"\n已輸出: {out_path}")


def main():
    db = get_db()
    existing = db.list_collection_names()

    while True:
        raw = input(
            "\n請輸入要查詢的 collection 名稱"
            "（all/export = 全部匯出 Excel；q = 離開）："
        ).strip()
        low = raw.lower()
        if low == "q":
            print("See you next time!")
            return
        if low in ("all", "export"):
            export_all(db)
            continue

        name = raw.upper()
        if name not in existing:
            print(f"找不到 collection：{name}")
            print(f"目前可用的 collections：{', '.join(sorted(existing))}")
            continue

        famids = fetch_famids(db[name])
        if not famids:
            print(f"\n[{name}] 該 collection 無 famid 資料。")
            continue

        print_stats(name, compute_stats(famids))


if __name__ == "__main__":
    main()
