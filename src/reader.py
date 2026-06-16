import pandas as pd
from config import SHARED_FIELDS


def _format_date(val):
    """格式化 record_date；無法解析或空值回傳 None（交由後續必要欄位檢查報錯）。"""
    if val is None or pd.isna(val) or str(val).strip() == "":
        return None
    try:
        return pd.to_datetime(val).strftime("%Y-%m-%d")
    except Exception:
        return None


def read_xlsx(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, sheet_name="import", dtype=str) # 只讀取 import sheet
    df.columns = df.columns.str.strip() # Remove leading/trailing whitespace from column names
    df = df.loc[:, df.columns.notna()]  # 過濾掉無效欄位名（如 NaN）

    # 狀況一：整欄必要欄位不存在 → 印出明確錯誤並中止
    missing_cols = [c for c in SHARED_FIELDS if c not in df.columns]
    if missing_cols:
        raise SystemExit(f"匯入中止：Excel 缺少必要欄位欄 {', '.join(missing_cols)}")

    df = df.dropna(how="all")  # 只過濾完全空白的列
    df = df.where(pd.notna(df), None)
    df["record_date"] = df["record_date"].apply(_format_date)
    df = df.reset_index(drop=True)
    return df
