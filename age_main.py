from dotenv import load_dotenv
import pandas as pd
from config import COLLECTION_MAP
from src.transformer import (
    parse_collection_config, get_rule, validate_and_convert,
)
from src.importer import get_db, upsert_many
from src.age_matcher import find_age_conflicts, estimate_record_date
from src.no_date_writer import write_no_date_xlsx
from src.error_writer import write_error_xlsx
from src.import_logger import write_import_log
from src.logger import log_summary
from src.utils.select_collections import select_collections
from src.utils.wait_and_retry import wait_and_retry

load_dotenv()

FILEPATH = "import_data.xlsx"


def read_xlsx_age(filepath: str) -> pd.DataFrame:
    """Read Excel, require famid + age (record_date not required)."""
    df = pd.read_excel(filepath, sheet_name="import", dtype=str)
    df.columns = df.columns.str.strip()
    df = df.where(pd.notna(df), None)
    df = df.dropna(subset=["famid", "age"])
    df = df.reset_index(drop=True)
    return df


def main():
    selected = select_collections()
    if selected is None:
        return

    print("reading file...")
    df = read_xlsx_age(FILEPATH)
    print(f"total rows (with age): {len(df)}")

    if df.empty:
        print("no rows with age found, exiting")
        return

    db = get_db()
    no_date_rows = []
    error_rows = []
    import_docs = {}   # collection -> list of docs to upsert
    skipped_count = 0

    items = [
        (k, v) for k, v in COLLECTION_MAP.items()
        if selected is None or k in selected
    ]

    for collection, fields in items:
        shared, prefixes, extra_col_names = parse_collection_config(fields)
        scale_cols = [
            col for col in df.columns
            if any(col.startswith(p) for p in prefixes)
        ]
        for col in extra_col_names:
            if col in df.columns and col not in scale_cols:
                scale_cols.append(col)

        col_no_date = []
        col_errors = 0
        col_import = []

        for _, row in df.iterrows():
            famid = row.get("famid")
            role = row.get("role")
            age_str = row.get("age")

            # Parse age
            try:
                raw_age = float(age_str)
            except (ValueError, TypeError):
                continue

            # Validate scale values
            errors = []
            converted = {}
            for col in scale_cols:
                val = row.get(col)
                rule = get_rule(collection, col)
                new_val, err = validate_and_convert(col, val, rule, collection)
                if err:
                    errors.append(err)
                else:
                    converted[col] = new_val

            # Skip if all scale fields empty
            if all(converted.get(col) is None for col in scale_cols):
                skipped_count += 1
                continue

            if errors:
                error_row = row.to_dict()
                error_row["error"] = f"{famid}: " + ", ".join(errors)
                error_row["collection"] = collection
                error_rows.append(error_row)
                col_errors += 1
                continue

            # Add metadata
            converted["famid"] = famid
            converted["role"] = role
            age_field = f"raw_{collection.lower()}_age"
            converted[age_field] = raw_age

            # Check for age conflicts with existing records
            matches, err_msg = find_age_conflicts(
                db, collection, famid, raw_age,
            )

            if err_msg:
                # Can't look up birth_date — flag for review
                col_no_date.append({
                    "collection": collection,
                    "reason": err_msg,
                    **converted,
                })
            elif matches:
                # Age within 1 year of existing record — potential duplicate
                for match in matches:
                    col_no_date.append({
                        "collection": collection,
                        "reason": (
                            f"與已存在紀錄年齡相差 {match['age_diff']} 歲 "
                            f"(record_date: {match['record_date']}, "
                            f"calculated_age: {match['calculated_age']})"
                        ),
                        **converted,
                    })
            else:
                # No conflict — estimate record_date and import
                est_date, est_err = estimate_record_date(db, famid, raw_age)
                if est_err:
                    col_no_date.append({
                        "collection": collection,
                        "reason": est_err,
                        **converted,
                    })
                else:
                    converted["record_date"] = est_date
                    converted["est_record_date"] = 1
                    col_import.append(converted)

        no_date_rows.extend(col_no_date)
        import_docs[collection] = col_import
        print(
            f"[{collection}] "
            f"to_import: {len(col_import)}, "
            f"age_conflicts: {len(col_no_date)}, "
            f"errors: {col_errors}"
        )

    # Write errors
    if error_rows:
        print("writing errors...")
        wait_and_retry(lambda: write_error_xlsx(error_rows), "errors.xlsx")

    # Write no_date_records
    if no_date_rows:
        print("writing no_date_records...")
        wait_and_retry(
            lambda: write_no_date_xlsx(no_date_rows),
            "no_date_records.xlsx",
        )

    # Upsert to MongoDB
    print("writing to MongoDB...")
    log_entries = []
    for col_name, docs in import_docs.items():
        error_count = len([r for r in error_rows if r.get("collection") == col_name])
        no_date_count = len([r for r in no_date_rows if r.get("collection") == col_name])
        skipped_col = len(df) - len(docs) - error_count - no_date_count
        log_summary(col_name, len(docs), skipped_col, error_count)
        upsert_many(db, col_name, docs)
        log_entries.append({
            "collection": col_name,
            "total": len(df),
            "success": len(docs),
            "errors": error_count,
            "skipped": skipped_col,
            "no_date_conflicts": no_date_count,
        })

    print("writing import log...")
    wait_and_retry(lambda: write_import_log(log_entries), "import_log.xlsx")

    print("done")


if __name__ == "__main__":
    main()
