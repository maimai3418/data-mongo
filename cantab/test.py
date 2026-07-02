import pandas as pd
from pathlib import Path

folder = Path(r"C:\Users\Dell_Insp_3030s\Desktop\20260107\CANTAB\new_cantab(明坊)\work")
files = list(folder.glob("*.xlsx"))
print(f"共找到 {len(files)} 個檔案")

for f in files:
    df = pd.read_excel(f, nrows=0)
    bad = [c for c in df.columns if not isinstance(c, str)]
    print(f.name, "→", bad if bad else "無異常")