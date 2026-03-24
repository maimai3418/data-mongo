"""驗證 config.py 中 COLLECTION_MAP 與 VALID_RANGE 的對應關係。"""
from config import COLLECTION_MAP, VALID_RANGE


def _has_own_rules(fields):
    """dict 格式且自帶 default_rule 的 collection 不需要 VALID_RANGE。"""
    return isinstance(fields, dict) and "default_rule" in fields


def test_collection_map_keys_have_valid_range():
    """每個 COLLECTION_MAP 的 key 都必須能在 get_rule 中解析到規則。
    - dict 格式且自帶 default_rule → OK
    - 否則必須存在於 VALID_RANGE
    """
    missing = []
    for name, fields in COLLECTION_MAP.items():
        if _has_own_rules(fields):
            continue
        if name not in VALID_RANGE:
            missing.append(name)
    assert missing == [], f"COLLECTION_MAP 中缺少 VALID_RANGE 對應: {missing}"


def test_valid_range_keys_in_collection_map():
    """VALID_RANGE 的每個 key 都應存在於 COLLECTION_MAP。"""
    extra = [k for k in VALID_RANGE if k not in COLLECTION_MAP]
    assert extra == [], f"VALID_RANGE 中有多餘的 key（不在 COLLECTION_MAP）: {extra}"


def test_collection_map_values_format():
    """COLLECTION_MAP 的值必須是 list 或 dict，且格式正確。"""
    for name, fields in COLLECTION_MAP.items():
        assert isinstance(fields, (list, dict)), (
            f"{name}: 值必須是 list 或 dict，實際為 {type(fields)}"
        )
        if isinstance(fields, dict):
            assert "prefixes" in fields, f"{name}: dict 格式缺少 'prefixes'"


def test_valid_range_rule_format():
    """VALID_RANGE 每個規則都必須有 type，且 type=int 時需有 range。"""
    for name, rule in VALID_RANGE.items():
        assert "type" in rule, f"{name}: 缺少 'type'"
        if rule["type"] == "int":
            assert "range" in rule, f"{name}: type=int 但缺少 'range'"
            lo, hi = rule["range"]
            assert lo <= hi, f"{name}: range 不合理 ({lo}, {hi})"


if __name__ == "__main__":
    tests = [
        test_collection_map_keys_have_valid_range,
        test_valid_range_keys_in_collection_map,
        test_collection_map_values_format,
        test_valid_range_rule_format,
    ]
    for t in tests:
        try:
            t()
            print(f"  PASS: {t.__name__}")
        except AssertionError as e:
            print(f"  FAIL: {t.__name__} -> {e}")
