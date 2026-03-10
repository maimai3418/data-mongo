SHARED_FIELDS = ["role", "famid",  "record_date"]

COLLECTION_MAP = {
    "GSQ":   SHARED_FIELDS + ["gsq_"],
    "SNAP4": SHARED_FIELDS + ["snap4_"],
    "CBCL":  SHARED_FIELDS + ["cbcl_"],
    "AQ":    SHARED_FIELDS + ["aq_"],
    "SRS":   SHARED_FIELDS + ["srs_"],
    "SAICA": SHARED_FIELDS + ["saica_"],
}

UNIQUE_KEY = ["famid", "record_date", "role"]

# 統一範圍的 collection
VALID_RANGE = {
    "GSQ":   {"type": "int", "range": (0, 6)},
    "SNAP4": {"type": "int", "range": (0, 3)},
    "CBCL":  {"type": "int", "range": (0, 3)},
    "AQ":    {"type": "int", "range": (0, 3)},
    "SRS":   {"type": "int", "range": (1, 4)},
    "SAICA": {"type": "per_field"},
}

# SAICA 個別欄位規則，其餘欄位套用 SAICA_DEFAULT_RULE
SAICA_FIELD_RULES = {
    "saica_b2_a": {"type": "float"},
    "saica_b2_b": {"type": "float"},
    "saica_b2_d": {"type": "float"},
    "saica_b2_c": {"type": "int", "range": (1, 3)},
}
SAICA_DEFAULT_RULE = {"type": "int", "range": (1, 4)}