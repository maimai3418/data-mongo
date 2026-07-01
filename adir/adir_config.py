# -*- coding: utf-8 -*-
"""
ADI-R 匯入設定。

與 ADOS（ados_config.py）的主要差異：
  1. 單一 collection（ADIR），不分 module。
  2. null sentinel（999）→ null，與 ADOS 相同；但「月齡欄位」例外：
     adir002, 005–010, 017, 019, 026, 028, 087 的 field.json range 為
     0–999，其 999 為合法值，先照 range 驗證，不轉 null。
     判定完全由 field.json 的 range 是否涵蓋 999 決定（不需維護欄位清單）。
  3. 工作表名稱用於自動標註 version（full / current），
     資料匯入同一 collection。
  4. Full 與 Current 為超集關係（共用 adirc* 欄位），不拆 collection。

欄位限制（型別 / 範圍）讀自 adir/fields 下的 field.json，
格式同 ADOS：
  - <INT>[lo,hi] / <FLOAT>[lo,hi] → 型別 + 範圍驗證
  - <INT> / <FLOAT>               → 只驗型別（無範圍）
  - <DATE>                        → 統一轉成 YYYY-MM-DD 文字
  - <string> / 其他               → 不轉換（原樣上傳）
"""

import os
import re
import json

# ── 基本設定 ────────────────────────────────────────────────────────────

ADIR_CURRENT_FIELDS = [
    "adire11", "adire12", "adire13", "adire14", "adire15", "adire16",
    "adir017", "adire18", "adir019", "adire20", "adire21", "adire22",
    "adire23", "adire24", "adire25", "adir026", "adire27", "adir028",
    "adirc29", "adirc30", "adirc31", "adirc32", "adirc33", "adirc34",
    "adirc35", "adirc36", "adirc37", "adirc38", "adirc39", "adirc40",
    "adirc41", "adirc42", "adirc43", "adirc44", "adirc45", "adirc46",
    "adirc47", "adirc48", "adirc49", "adirc50", "adirc51", "adirc52",
    "adirc53", "adirc54", "adirc55", "adirc56", "adirc57", "adirc58",
    "adirc59", "adirc60", "adirc61", "adirc62", "adirc63", "adirc64",
    "adirc65", "adira65", "adirc66", "adirc67", "adirc68", "adirc69",
    "adirc70", "adirc71", "adirc72", "adirc73", "adirc74", "adirc75",
    "adirc76", "adirc77", "adirc78", "adirc79", "adirc80", "adirc81",
    "adirc82", "adirc83", "adirc84", "adirc85", "adirc88", "adirc89",
    "adirc90", "adirc91", "adirc92", "adirc93",
]




# 單一 collection（Full / Current 共存，以 version 欄位區分）
ADIR_COLLECTION = "ADIR"

# 重複判定用的 unique key：相同 famid + record_date 視為重複
ADIR_SHARED_FIELDS = ["famid", "record_date"]

# null sentinel：item 欄位等於此值 → null（遺漏值），與 ADOS 相同。
# 例外：月齡欄位（field.json range 涵蓋 999，如 0–999）的 999 為合法值，不轉 null。
ADIR_NULL_VALUE = 999

# header 欄位：不參與「item 全空→略過」判定，清空白 + 日期正規化即可
ADIR_HEADER_FIELDS = [
    "famid", "record_date", "birth_date", "sex",
    "p_rel", "p_rel_text", "interviewer", "interview_location",
    "improvement", "employment", "grade", "version",
]

# 工作表名稱（小寫）→ version 自動標註
# 若工作表名不在此表中，version 由資料本身或 CLI 參數決定
ADIR_SHEET_VERSION_MAP = {
    "full":         "full",
    "current":      "current",
    "adir full":    "full",
    "adir current": "current",
    "adir_full":    "full",
    "adir_current": "current",
}

# field.json 所在資料夾（與本檔同層的 fields/）
ADIR_FIELDS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fields")

# 臨時覆寫規則：{欄位: rule}，優先於 field.json（預設為空）
ADIR_FIELD_RULE_OVERRIDES = {}



# ── field.json 解析 ────────────────────────────────────────────────────

_SPEC_RE = re.compile(
    r"^<(?P<type>[A-Za-z]+)>(?:\[\s*(?P<lo>[^,\]]+)\s*,\s*(?P<hi>[^,\]]+)\s*\])?$"
)


def _to_num(s):
    """字串轉 int（整數值）或 float。"""
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
    """找出 ADI-R 的 field.json：檔名（小寫）含 'adir' 且以 'fields.json' 結尾；
    多個取檔名最大者（日期前綴最新）。"""
    fields_dir = fields_dir or ADIR_FIELDS_DIR
    if not os.path.isdir(fields_dir):
        return None
    matches = sorted(
        f for f in os.listdir(fields_dir)
        if f.lower().endswith("fields.json") and "adir" in f.lower()
    )
    return os.path.join(fields_dir, matches[-1]) if matches else None


def load_field_rules(fields_dir=None):
    """讀取 field.json，回傳 {欄位: rule}（已套用 OVERRIDES）。

    與 ADOS 不同：回傳扁平 dict（無 module 層級）。
    """
    fields_dir = fields_dir or ADIR_FIELDS_DIR
    rules = {}
    path = discover_field_json(fields_dir)
    if path:
        with open(path, "r", encoding="utf-8") as f:
            fields = json.load(f)
        rules = {col: parse_field_spec(spec) for col, spec in fields.items()}
    rules.update(ADIR_FIELD_RULE_OVERRIDES)
    return rules