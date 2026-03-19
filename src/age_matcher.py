from datetime import datetime, timedelta


def parse_date(val):
    """Parse date string or datetime object to datetime."""
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
    return None


def calculate_age(birth_date, record_date):
    """Calculate age in years from birth_date and record_date."""
    b = parse_date(birth_date)
    r = parse_date(record_date)
    if not b or not r:
        return None
    return (r - b).days / 365.25


def find_age_conflicts(db, collection_name, famid, raw_age):
    """
    Check if famid has existing records in collection
    whose calculated age (from Participants birth_date + record_date)
    is within 1 year of raw_age.

    Returns:
        (matches, error_msg)
        - matches: list of dicts with record_date, calculated_age, age_diff
        - error_msg: str if lookup failed, else None
    """
    participant = db["Participants"].find_one({"famid": famid})
    if not participant:
        return None, "Participants 中找不到此 famid"

    birth_date = participant.get("birth_date")
    if not birth_date:
        return None, "Participants 中無生日資料"

    existing = list(db[collection_name].find({"famid": famid}))
    if not existing:
        return [], None

    matches = []
    for record in existing:
        record_date = record.get("record_date")
        if not record_date:
            continue
        age = calculate_age(birth_date, record_date)
        if age is None:
            continue
        diff = abs(age - raw_age)
        if diff <= 1:
            matches.append({
                "record_date": record_date,
                "calculated_age": round(age, 2),
                "age_diff": round(diff, 2),
            })

    return matches, None


def estimate_record_date(db, famid, raw_age):
    """
    Estimate record_date from Participants birth_date + age (years).
    Returns (date_str, error_msg).
    """
    participant = db["Participants"].find_one({"famid": famid})
    if not participant:
        return None, "Participants 中找不到此 famid"

    birth_date = participant.get("birth_date")
    if not birth_date:
        return None, "Participants 中無生日資料"

    b = parse_date(birth_date)
    if not b:
        return None, "birth_date 格式錯誤"

    est = b + timedelta(days=raw_age * 365.25)
    return est.strftime("%Y-%m-%d"), None
