import pandas as pd

def read_xlsx(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, dtype=str)
    df.columns = df.columns.str.strip()

    if "record_date" in df.columns:
        df["record_date"] = pd.to_datetime(df["record_date"]).dt.strftime("%Y-%m-%d")

    df = df.where(pd.notna(df), None)
    return df