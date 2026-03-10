import pandas as pd
from config import COLLECTION_MAP, SHARED_FIELDS, VALID_RANGE, SAICA_FIELD_RULES, SAICA_DEFAULT_RULE


def validate_and_convert(col, val, rule, collection):
    if val is None:
        return None, None

    # 999 跳過驗證，直接存 null
    if str(val).strip() == "999":
        return None, None

    if rule["type"] == "float":
        try:
            return float(val), None
        except (ValueError, TypeError):
            return None, f"{col} 數值錯誤"

    if rule["type"] == "int":
        try:
            int_val = int(float(val))
        except (ValueError, TypeError):
            return None, f"{col} 數值錯誤"

        # if collection == "AQ":
        #     int_val = int_val - 1

        if "range" in rule:
            min_val, max_val = rule["range"]
            if not (min_val <= int_val <= max_val):
                return None, f"{col} 數值錯誤"

        return int_val, None


def get_rule(collection, col):
    if collection == "SAICA":
        return SAICA_FIELD_RULES.get(col, SAICA_DEFAULT_RULE)
    return VALID_RANGE[collection]


def split_by_collection(df: pd.DataFrame) -> tuple:
    valid_result = {}
    error_rows = []

    for collection, fields in COLLECTION_MAP.items():
        prefixes = [f for f in fields if f not in SHARED_FIELDS]
        scale_cols = [
            col for col in df.columns
            if any(col.startswith(p) for p in prefixes)
        ]
        existing_cols = [c for c in SHARED_FIELDS + scale_cols if c in df.columns]
        sub_df = df[existing_cols].copy()

        valid_docs = []

        for _, row in sub_df.iterrows():
            errors = []
            converted = row.to_dict()

            for col in scale_cols:
                val = converted.get(col)
                rule = get_rule(collection, col)
                new_val, err = validate_and_convert(col, val, rule, collection)
                if err:
                    errors.append(err)
                else:
                    converted[col] = new_val

            if errors:
                error_row = row.to_dict()
                famid = error_row.get("famid", "")
                record_date = error_row.get("record_date", "")
                error_row["error"] = f"{famid}-{record_date}: " + ", ".join(errors)
                error_row["collection"] = collection
                error_rows.append(error_row)
            else:
                valid_docs.append(converted)

        valid_result[collection] = valid_docs

    return valid_result, error_rows