import pandas as pd

def read_xlsx(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, sheet_name="import", dtype=str) # 只讀取 import sheet
    df.columns = df.columns.str.strip() # Remove leading/trailing whitespace from column names
    df = df.loc[:, df.columns.notna()]  # 過濾掉無效欄位名（如 NaN）

    if "record_date" in df.columns:
        df["record_date"] = pd.to_datetime(df["record_date"]).dt.strftime("%Y-%m-%d")

    df = df.where(pd.notna(df), None)
    # print(df.columns.tolist())
    df = df.dropna(subset=["famid", "record_date"])  # 過濾空白行
    df = df.reset_index(drop=True)
    return df