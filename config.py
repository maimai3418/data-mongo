SHARED_FIELDS = ["role", "famid",  "record_date"]

COLLECTION_MAP = {
    "GSQ":   SHARED_FIELDS + ["gsq_"],
    "SNAP4": SHARED_FIELDS + ["snap4_"],
    "CBCL":  SHARED_FIELDS + ["cbcl_"],
    "AQ":    SHARED_FIELDS + ["aq_"],
    "SRS":   SHARED_FIELDS + ["srs_"],
    "SAICA": SHARED_FIELDS + ["saica_"],
    "SCQ":   SHARED_FIELDS + ["scq_"],
    "CAST":  SHARED_FIELDS + ["cast_"],
    "BRIEF": SHARED_FIELDS + ["brief_"],
    "CPRS":   SHARED_FIELDS + ["cprs_"],
    "CTRS":   SHARED_FIELDS + ["ctrs_"],
    "SDQ":    SHARED_FIELDS + ["sdq_"],
    "PMI":    SHARED_FIELDS + ["pmi_"],
    "MPBI":    SHARED_FIELDS + ["mpbi_"],
    "DPBI":    SHARED_FIELDS + ["dpbi_"],
    "APGAR":  SHARED_FIELDS + ["apgar_"],
    "EPQ":    SHARED_FIELDS + ["epq_"],
    "YSR":    SHARED_FIELDS + ["ysr_"],
    # "FACES":  SHARED_FIELDS + ["faces_"],
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
    "SCQ":   {"type": "int", "range": (0, 1)},
    "CAST":  {"type": "int", "range": (0, 1)},
    "BRIEF": {"type": "int", "range": (1, 3)},
    "CPRS":   {"type": "int", "range": (0, 3)},
    "CTRS":   {"type": "int", "range": (0, 3)},
    "SDQ":     {"type": "per_field"},
    "PMI":    {"type": "int", "range": (1, 4)},
    "MPBI":    {"type": "int", "range": (1, 4)},
    "DPBI":    {"type": "int", "range": (1, 4)},
    "APGAR":  {"type": "int", "range": (0, 2)},
    "EPQ":    {"type": "int", "range": (0, 1)},
    "YSR":    {"type": "int", "range": (0, 2)},
}

SDQ_FIELD_RULES = {
    "sdq_26": {"type": "int", "range": (0, 3)},
    "sdq_27": {"type": "int", "range": (0, 3)},
    "sdq_28": {"type": "int", "range": (0, 3)},
    "sdq_29_a": {"type": "int", "range": (0, 3)},
    "sdq_29_b": {"type": "int", "range": (0, 3)},
    "sdq_29_c": {"type": "int", "range": (0, 3)},
    "sdq_29_d": {"type": "int", "range": (0, 3)},
    "sdq_30": {"type": "int", "range": (0, 3)},
}
SDQ_DEFAULT_RULE = {"type": "int", "range": (0, 2)}

# SAICA 個別欄位規則，其餘欄位套用 SAICA_DEFAULT_RULE
SAICA_FIELD_RULES = {
    "saica_b2_a": {"type": "float"},
    "saica_b2_b": {"type": "float"},
    "saica_b2_d": {"type": "float"},
    "saica_b2_c": {"type": "int", "range": (1, 3)},
}
SAICA_DEFAULT_RULE = {"type": "int", "range": (1, 4)}