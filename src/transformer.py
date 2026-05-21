import pandas as pd
from config import COLLECTION_MAP, SHARED_FIELDS, VALID_RANGE
from config import SAICA_FIELD_RULES, SAICA_DEFAULT_RULE
from config import SDQ_FIELD_RULES, SDQ_DEFAULT_RULE
from config import SF36_FIELD_RULES, SF36_DEFAULT_RULE


def validate_and_convert(col, val, rule, collection):
    if val is None or pd.isna(val) or str(val).strip() in ("", "999", "N/A", "n/a", "NA", "na"):
        return None, None

    if rule["type"] == "float":
        try:
            float_val = float(val)
        except (ValueError, TypeError):
            return None, f"{col} 數值錯誤"
        if "min" in rule and float_val < rule["min"]:
            return None, f"{col} 數值錯誤"
        if "max" in rule and float_val > rule["max"]:
            return None, f"{col} 數值錯誤"
        return float_val, None

    if rule["type"] == "int":
        try:
            float_val = float(val)
            if float_val != int(float_val):  # 2.5 != 2 -> 跳錯
                return None, f"{col} 數值錯誤(非整數)"
            int_val = int(float_val)
        except (ValueError, TypeError):
            return None, f"{col} 數值錯誤"

        # # AQ 轉換 4->3, 3->2, 2->1, 1->0
        # # MOST109 ~ NHRI110 早期問卷用的 AQ 計分(prc不變)要改為 0~3 分
        # # 原本的 1~4 分改為 0~3 分（即原本的 4 分變成 3 分，3 分變成 2 分，以此類推）
        if collection == "AQ":
            int_val = int_val - 1

        if "range" in rule:
            min_val, max_val = rule["range"]
            if not (min_val <= int_val <= max_val):
                return None, f"{col} 數值錯誤"

        return int_val, None


def get_rule(collection, col):
    fields = COLLECTION_MAP.get(collection)
    if isinstance(fields, dict):
        # extra_cols 自帶規則
        extra_cols = fields.get("extra_cols", {})
        if col in extra_cols:
            return extra_cols[col]
        # prefix 欄位：先查 field_rules，再用 default_rule
        field_rules = fields.get("field_rules", {})
        if col in field_rules:
            return field_rules[col]
        if "default_rule" in fields:
            return fields["default_rule"]
    # list 格式的 per_field 量表
    if collection == "SAICA":
        return SAICA_FIELD_RULES.get(col, SAICA_DEFAULT_RULE)
    if collection == "SDQ":
        return SDQ_FIELD_RULES.get(col, SDQ_DEFAULT_RULE)
    if collection == "SF-36":
        return SF36_FIELD_RULES.get(col, SF36_DEFAULT_RULE)
    return VALID_RANGE[collection]


def parse_collection_config(fields):
    """解析 COLLECTION_MAP 的值，支援 list 和 dict 兩種格式。"""
    if isinstance(fields, dict):
        prefixes = fields.get("prefixes", [])
        extra_cols = fields.get("extra_cols", {})
        # extra_cols 可以是 dict (帶規則) 或 list (僅名稱)
        extra_col_names = list(extra_cols.keys()) if isinstance(extra_cols, dict) else list(extra_cols)
        shared = fields.get("shared_fields", SHARED_FIELDS)
        return shared, prefixes, extra_col_names
    else:
        prefixes = [f for f in fields if f not in SHARED_FIELDS]
        return list(SHARED_FIELDS), prefixes, []


def split_by_collection(df: pd.DataFrame, selected=None) -> tuple:
    valid_result = {}
    error_rows = []
    skipped_rows = []

    items = [(k, v) for k, v in COLLECTION_MAP.items() if selected is None or k in selected]
    for collection, fields in items:
        shared, prefixes, extra_cols = parse_collection_config(fields)
        scale_cols = [
            col for col in df.columns
            if any(col.startswith(p) for p in prefixes)
        ]
        # 加入額外指定的欄位（存在於 df 中的）
        for col in extra_cols:
            if col in df.columns and col not in scale_cols:
                scale_cols.append(col)
        existing_cols = [c for c in shared + scale_cols if c in df.columns]
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

            # 量表欄位全為空值，跳過不上傳
            if all(converted.get(col) is None for col in scale_cols):
                skipped_rows.append({"collection": collection, **row.to_dict()})
                continue

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

    return valid_result, error_rows, skipped_rows