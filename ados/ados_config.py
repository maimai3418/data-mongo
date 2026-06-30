# -*- coding: utf-8 -*-
"""
ADOS 匯入設定。

與一般 scale 量表（config.py）的差異：
  1. 重複判定 unique key 只看 famid + record_date（不看 role）。
  2. 數值轉換：只有等於 ADOS_NULL_VALUE（999）才轉成 null，
     其餘一律按原始值上傳（0 傳 0、1 傳 1…）。
  3. 工作表名稱直接對應 collection：M1 → ADOS_M1、M2 → ADOS_M2…

欄位限制（型別 / 範圍）**直接讀自 ados/fields 下的 field.json**，
格式如 "<INT>[0,8]"、"<INT>[1,2]"、"<DATE>"、"<string>"：
  - <INT>[lo,hi] / <FLOAT>[lo,hi] → 套用型別 + 範圍驗證
  - <INT> / <FLOAT>               → 只驗證型別（無範圍）
  - <DATE>                        → 統一轉成 YYYY-MM-DD 文字（無法解析則報錯）
  - <string> / 其他               → 不轉換（原樣上傳）
匯入時對 item 欄位（ados_<module>_*）套用數值限制；對 <DATE> 欄位做日期正規化。
若要臨時覆寫某欄位規則，填入 ADOS_FIELD_RULE_OVERRIDES（優先於 field.json）。
"""

import os
import re
import json

# 重複判定用的 unique key：相同 famid + record_date 視為重複，不重複匯入（不看 role）
ADOS_SHARED_FIELDS = ["famid", "record_date"]

# 工作表名稱 → MongoDB collection 名稱
ADOS_COLLECTION_MAP = {
    "M1": "ADOS_M1",
    "M2": "ADOS_M2",
    "M3": "ADOS_M3",
    "M4": "ADOS_M4",
}

# 只有等於此值才轉成 null；其餘維持原始值上傳
ADOS_NULL_VALUE = 999

# header 欄位：原樣上傳（清掉前後空白、空字串→null），不套用數值範圍驗證
ADOS_HEADER_FIELDS = ["famid", "record_date", "birth_date", "sex", "interviewer"]

# item 欄位前綴：ados_<module 小寫>_，例如 M1 → "ados_m1_"
ADOS_ITEM_PREFIX = "ados_{module}_"

# field.json 所在資料夾（與本檔同層的 fields/）
ADOS_FIELDS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fields")

# 臨時覆寫規則：{module: {欄位: rule}}，優先於 field.json 解析結果（預設為空）。
# rule 格式同解析結果：{"type": "int"/"float", "range": (min, max)}
ADOS_FIELD_RULE_OVERRIDES = {}


# field.json 型別字串格式，如 "<INT>[0,8]"、"<INT>"、"<DATE>"、"<string>"
_SPEC_RE = re.compile(
    r"^<(?P<type>[A-Za-z]+)>(?:\[\s*(?P<lo>[^,\]]+)\s*,\s*(?P<hi>[^,\]]+)\s*\])?$"
)


def _to_num(s):
    """字串轉 int（整數值）或 float。"""
    f = float(s)
    return int(f) if f.is_integer() else f


def parse_field_spec(spec):
    """把 field.json 的型別字串解析成驗證規則。

    "<INT>[0,8]" → {"type":"int","range":(0,8)}；"<INT>" → {"type":"int"}；
    "<FLOAT>…"   → float；其餘（<string>/<DATE>/無法解析）→ None（原樣上傳、不驗證）。
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
        return {"type": "date"}  # 匯入時統一轉成 YYYY-MM-DD 文字
    else:
        return None  # string 等：不做轉換 / 驗證
    lo, hi = m.group("lo"), m.group("hi")
    if lo is not None and hi is not None:
        try:
            rule["range"] = (_to_num(lo), _to_num(hi))
        except (ValueError, TypeError):
            pass  # 範圍無法解析時，僅保留型別驗證
    return rule


def discover_field_json(fields_dir, module):
    """找出該 module 的 field.json：檔名（小寫）含 'ados_<module>' 且以 'fields.json' 結尾；
    多個取檔名最大者（日期前綴最新）。找不到回 None。"""
    if not os.path.isdir(fields_dir):
        return None
    needle = f"ados_{module.lower()}"
    matches = sorted(
        f for f in os.listdir(fields_dir)
        if f.lower().endswith("fields.json") and needle in f.lower()
    )
    return os.path.join(fields_dir, matches[-1]) if matches else None


def load_field_rules(fields_dir=None):
    """讀取各 module 的 field.json，回傳 {module: {欄位: rule}}（已套用 OVERRIDES）。

    rule 為 parse_field_spec 的結果（數值欄位為 dict，其餘為 None）。
    """
    fields_dir = fields_dir or ADOS_FIELDS_DIR
    rules = {}
    for module in ADOS_COLLECTION_MAP:
        module_rules = {}
        path = discover_field_json(fields_dir, module)
        if path:
            with open(path, "r", encoding="utf-8") as f:
                fields = json.load(f)
            module_rules = {col: parse_field_spec(spec) for col, spec in fields.items()}
        # 套用手動覆寫（優先於 field.json）
        module_rules.update(ADOS_FIELD_RULE_OVERRIDES.get(module, {}))
        rules[module] = module_rules
    return rules
