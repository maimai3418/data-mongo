# -*- coding: utf-8 -*-
"""
CPT 匯入設定。

與 ADI-R（adir_config.py）同款：欄位限制讀自 cpt/fields 下的 field.json，
格式：
  - <INT>[lo,hi] / <FLOAT>[lo,hi] → 型別 + 範圍驗證
  - <INT> / <FLOAT>               → 只驗型別（無範圍）
  - <DATE>                        → 統一轉成 YYYY-MM-DD 文字
  - <string> / 其他               → 不轉換（原樣上傳）

CPT 沒有 999 null sentinel 慣例（field.json 未見任何 range 涵蓋 999 的設計），
不做全域 999→null 轉換。
"""

import os
import re
import json

# 單一 collection
CPT_COLLECTION = "CPT"

# 重複判定 / 三層去重複用的 unique key：famid 在 CPT 語境直接等同個人 ID
CPT_SHARED_FIELDS = ["famid", "record_date"]

# header（metadata）欄位：不參與去重複時的欄位值比對，只做清理 + 型別轉換
CPT_HEADER_FIELDS = ["famid", "sex", "birth_date", "record_date"]

# field.json 所在資料夾（與本檔同層的 fields/）
CPT_FIELDS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fields")

# 臨時覆寫規則：{欄位: rule}，優先於 field.json（預設為空）
CPT_FIELD_RULE_OVERRIDES = {}


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
    """找出 CPT 的 field.json：檔名（小寫）含 'cpt' 且以 'fields.json' 結尾；
    多個取檔名最大者（日期前綴最新）。"""
    fields_dir = fields_dir or CPT_FIELDS_DIR
    if not os.path.isdir(fields_dir):
        return None
    matches = sorted(
        f for f in os.listdir(fields_dir)
        if f.lower().endswith("fields.json") and "cpt" in f.lower()
    )
    return os.path.join(fields_dir, matches[-1]) if matches else None


def load_field_rules(fields_dir=None):
    """讀取 field.json，回傳 {欄位: rule}（已套用 OVERRIDES）。"""
    fields_dir = fields_dir or CPT_FIELDS_DIR
    rules = {}
    path = discover_field_json(fields_dir)
    if path:
        with open(path, "r", encoding="utf-8") as f:
            fields = json.load(f)
        rules = {col: parse_field_spec(spec) for col, spec in fields.items()}
    rules.update(CPT_FIELD_RULE_OVERRIDES)
    return rules
