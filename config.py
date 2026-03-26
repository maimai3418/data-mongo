SHARED_FIELDS = ["role", "famid",  "record_date"]

# ASRI 額外的子量表分數欄位（非 asri_ 開頭）
# key: 欄位名, value: 該欄位的驗證規則
_ASRI_EXTRA_RULE = {"type": "float", "min": 0}
ASRI_EXTRA_COLS = {
    "agad_ag": _ASRI_EXTRA_RULE,
    "aspepho_ag": _ASRI_EXTRA_RULE,
    "aagora_ag": _ASRI_EXTRA_RULE,
    "apanic_ag": _ASRI_EXTRA_RULE,
    "aobs_ag": _ASRI_EXTRA_RULE,
    "acom_ag": _ASRI_EXTRA_RULE,
    "amotic_ag": _ASRI_EXTRA_RULE,
    "avotic_ag": _ASRI_EXTRA_RULE,
    "asocpho_ag": _ASRI_EXTRA_RULE,
    "aschizoid_ag": _ASRI_EXTRA_RULE,
    "asomat_ag": _ASRI_EXTRA_RULE,
    "ahypoch_ag": _ASRI_EXTRA_RULE,
    "adysmor_ag": _ASRI_EXTRA_RULE,
    "agendid_ag": _ASRI_EXTRA_RULE,
    "aanorex_ag": _ASRI_EXTRA_RULE,
    "abulimia_ag": _ASRI_EXTRA_RULE,
    "anarcol_ag": _ASRI_EXTRA_RULE,
    "anightmar_ag": _ASRI_EXTRA_RULE,
    "aodd_ag": _ASRI_EXTRA_RULE,
    "acd_ag": _ASRI_EXTRA_RULE,
    "aantiso_ag": _ASRI_EXTRA_RULE,
    "aschizo_ag": _ASRI_EXTRA_RULE,
    "asub_ag": _ASRI_EXTRA_RULE,
    "adissoc_ag": _ASRI_EXTRA_RULE,
    "adepress_ag": _ASRI_EXTRA_RULE,
    "aadjust_ag": _ASRI_EXTRA_RULE,
    "abpd_ag": _ASRI_EXTRA_RULE,
    "agambli_ag": _ASRI_EXTRA_RULE,
    "atrichot_ag": _ASRI_EXTRA_RULE,
}

ASRI_FIELD_RULES = {
    "asri_48":{"type": "int", "range": (0, 1)},
    "asri_49":{"type": "int", "range": (0, 1)},
    "asri_50":{"type": "int", "range": (0, 1)},
}
ASRI_DEFAULT_RULE = {"type": "int", "range": (0, 3)}

# COLLECTION_MAP 支援兩種格式：
#   list: SHARED_FIELDS + ["prefix_"]  （向下相容）
#   dict: {"prefixes": [...], "extra_cols": {...}, "shared_fields": [...],
#          "field_rules": {...}, "default_rule": {...}}
#         - prefixes: 欄位前綴（必填）
#         - extra_cols: {欄位名: 驗證規則} （選填，預設 {}）
#         - shared_fields: 覆寫共用欄位（選填，預設用 SHARED_FIELDS）
#         - field_rules: {欄位名: 驗證規則} prefix 欄位的個別規則（選填）
#         - default_rule: prefix 欄位的預設規則（選填）
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
    "SUB":    SHARED_FIELDS + ["sub_"],
    "ASRI": {
        "prefixes": ["asri_"],
        "extra_cols": ASRI_EXTRA_COLS,
        "field_rules": ASRI_FIELD_RULES,
        "default_rule": ASRI_DEFAULT_RULE,
    },
    "ASRS":   SHARED_FIELDS + ["asrs_"],
    "CES-D":    SHARED_FIELDS + ["cesd_"],
    "MPI":    SHARED_FIELDS + ["mpi_"],
    "RBS-R":   SHARED_FIELDS + ["rbsr_"],
    "SSP":    SHARED_FIELDS + ["ssp_"],
    "WFIRS-P":   SHARED_FIELDS + ["weip_"],
    "WFIRS-S":   SHARED_FIELDS + ["weis_"],
    "WHOQOL":  SHARED_FIELDS + ["who_"],
    "CEQ":    SHARED_FIELDS + ["ceq_"],
    "C-SBEQ": SHARED_FIELDS + ["csbeq_"],
    "ERQ_CA": SHARED_FIELDS + ["erqc_"],
    "BRIEF-S": SHARED_FIELDS + ["briefs_"],
    "BRIEF-A": SHARED_FIELDS + ["briefa_"],
    "ERQ_A": SHARED_FIELDS + ["erqa_"],
    "RAADS-R": SHARED_FIELDS + ["raadsr_"],
    "TAS-20": SHARED_FIELDS + ["tas_"],
    "AAQOL": SHARED_FIELDS + ["aaqol_"],
    "ARI": SHARED_FIELDS + ["ari_"],
    "ESQ": SHARED_FIELDS + ["esq_"],
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
    "SUB":    {"type": "int", "range": (1, 4)},
    "ASRS":   {"type": "int", "range": (0, 4)},
    "CES-D":  {"type": "int", "range": (0, 3)},
    "MPI":    {"type": "int", "range": (0, 1)},
    "RBS-R":   {"type": "int", "range": (0, 3)},
    "SSP":    {"type": "int", "range": (1, 5)},
    "WFIRS-P":   {"type": "int", "range": (0, 3)},
    "WFIRS-S":   {"type": "int", "range": (0, 3)},
    "WHOQOL":  {"type": "int", "range": (0, 4)},
    "CEQ":    {"type": "int", "range": (0, 3)},
    "C-SBEQ": {"type": "int", "range": (0, 3)},
    "ERQ_CA": {"type": "int", "range": (1, 5)},
    "ERQ_A": {"type": "int", "range": (1, 7)},
    "RAADS-R": {"type": "int", "range": (0, 3)},
    "TAS-20": {"type": "int", "range": (1, 5)},
    "BRIEF-S": {"type": "int", "range": (1, 3)},
    "BRIEF-A": {"type": "int", "range": (0, 2)},
    "AAQOL": {"type": "int", "range": (1, 5)},
    "ARI": {"type": "int", "range": (0, 2)},
    "ESQ": {"type": "int", "range": (1, 4)},
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
    "sdq_31": {"type": "int", "range": (1, 5)},
    "sdq_32": {"type": "int", "range": (1, 5)},

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