# -*- coding: utf-8 -*-
"""
CANTAB 匯入設定。

與 CPT（cpt_config.py）同款：欄位限制讀自 cantab/fields 下的 field.json，
格式：
  - <INT>[lo,hi] / <FLOAT>[lo,hi] → 型別 + 範圍驗證
  - <INT> / <FLOAT>               → 只驗型別（無範圍）
  - <DATE>                        → 統一轉成 YYYY-MM-DD 文字
  - <string> / 其他               → 不轉換（原樣上傳）

CANTAB 沒有 999 null sentinel 慣例，不做全域 999→null 轉換。
field.json 的量表欄位名稱為原始大小寫（如 "MOTmL"、"PRMmcL"），比對時保留原樣，
不像 cantab_precheck.py 那樣統一轉小寫（轉小寫是 precheck 內部去重複比對用，
實際存入 DB 的欄位名稱仍需與 field.json 完全一致）。
"""

import os
import re
import json

# 單一 collection
CANTAB_COLLECTION = "CANTAB"

# 重複判定 / 三層去重複用的 unique key：famid 在 CANTAB 語境直接等同個人 ID
CANTAB_SHARED_FIELDS = ["famid", "record_date"]

# header（metadata）欄位：不參與去重複時的欄位值比對，只做清理 + 型別轉換
CANTAB_HEADER_FIELDS = ["famid", "sex", "birth_date", "record_date", "session_start_time"]

# 原始資料中可能出現的「年齡」欄位別名（大小寫不拘）；field.json 未定義此欄位，
# 只在完全沒有 record_date 時作為 age_suspect 判定與 raw_cantab_age 來源
AGE_ALIASES = ("age_cantab", "age")
RAW_AGE_FIELD = "raw_cantab_age"

# field.json 所在資料夾（與本檔同層的 fields/）
CANTAB_FIELDS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fields")

# 臨時覆寫規則：{欄位: rule}，優先於 field.json（預設為空）
CANTAB_FIELD_RULE_OVERRIDES = {}


_SPEC_RE = re.compile(
    r"^<(?P<type>[A-Za-z]+)>(?:\[\s*(?P<lo>[^,\]]+)\s*,\s*(?P<hi>[^,\]]+)\s*\])?$"
)


def _to_num(s):
    f = float(s)
    return int(f) if f.is_integer() else f


def parse_field_spec(spec):
    """把 field.json 的型別字串解析成驗證規則。

    "<INT>[0,9]" → {"type":"int","range":(0,9)}；"<INT>" → {"type":"int"}；
    "<DATE>"     → {"type":"date"}；其餘（<string> 等）→ None。
    """
    if not isinstance(spec, str):
        return None
    m = _SPEC_RE.match(spec.strip())
    if not m:
        return None
    t = m.group("type").lower()
    if t == "int":
        rule = {"type": "int"}
    elif t == "float":
        rule = {"type": "float"}
    elif t == "date":
        return {"type": "date"}
    else:
        return None
    lo, hi = m.group("lo"), m.group("hi")
    if lo is not None and hi is not None:
        try:
            rule["range"] = (_to_num(lo), _to_num(hi))
        except (ValueError, TypeError):
            pass
    return rule


def discover_field_json(fields_dir=None):
    """找出 CANTAB 的 field.json：檔名（小寫）含 'cantab' 且以 'fields.json' 結尾；
    多個取檔名最大者（日期前綴最新）。"""
    fields_dir = fields_dir or CANTAB_FIELDS_DIR
    if not os.path.isdir(fields_dir):
        return None
    matches = sorted(
        f for f in os.listdir(fields_dir)
        if f.lower().endswith("fields.json") and "cantab" in f.lower()
    )
    return os.path.join(fields_dir, matches[-1]) if matches else None


def load_field_rules(fields_dir=None):
    """讀取 field.json，回傳 {欄位: rule}（已套用 OVERRIDES）。欄位名稱保留原始大小寫。"""
    fields_dir = fields_dir or CANTAB_FIELDS_DIR
    rules = {}
    path = discover_field_json(fields_dir)
    if path:
        with open(path, "r", encoding="utf-8") as f:
            fields = json.load(f)
        rules = {col: parse_field_spec(spec) for col, spec in fields.items()}
    rules.update(CANTAB_FIELD_RULE_OVERRIDES)
    return rules
