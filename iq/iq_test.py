import pandas as pd



custom_path = "C:/Users/Dell_Insp_3030s/Desktop/20260107/IQ/new_IQ(明坊)"
filename = "20260702_iq.xlsx"
worksheet_name = "famid_check"

# df = pd.read_excel(f"{custom_path}/{filename}",
#                     sheet_name=worksheet_name,
#                     engine="openpyxl",
#                     dtype=str)  # famid 一律用字串讀取，避免數值型態產生 .0

# cols = ["famid_1", "famid_2", "famid_3", "famid_4"]
# unique_vals = pd.unique(df[cols].values.ravel())
# unique_vals = [v for v in unique_vals if pd.notna(v) and v != ""]

# # 若欄位混雜數值型態，清掉 Excel float artifact
# unique_vals = [v[:-2] if v.endswith(".0") else v for v in unique_vals]

# result_df = pd.DataFrame({"famid": sorted(unique_vals)})

# with pd.ExcelWriter(
#     f"{custom_path}/{filename}",
#     engine="openpyxl",
#     mode="a",                    # append，保留原有分頁
#     if_sheet_exists="replace"    # 同名分頁已存在時覆蓋；第一次執行可省略
# ) as writer:
#     result_df.to_excel(writer, sheet_name="famid_unique", index=False)



# ===== CONFIG =====

FILE_A = "C:/Users/Dell_Insp_3030s/Desktop/20260107/IQ/new_IQ(明坊)/work/MOST107IQ_200907_most107.xlsx"
FILE_B = "C:/Users/Dell_Insp_3030s/Desktop/20260107/IQ/new_IQ(明坊)/work/WPPSI221012_nsc96_asd.xlsx"

FAMID_COL = "famid"
COMPARE_COLS = ["FIQ", "PIQ", "VIQ"]

TARGET_FAMIDS = [
    "410301","410364","410494","410521","411151","411441","411661","411751",
    "411761","411791","411821","412051","412341","412394","412431","412434",
    "412681","412684","413694","413844","414201","414221","414391","414544",
    "414561","414944","415341","415711","415854","416031","416411","416731",
    "420911","420924","440294",
]
# ===================

def prep(path):
    df = pd.read_excel(path, engine="openpyxl", dtype=str)
    df[FAMID_COL] = df[FAMID_COL].str.replace(r"\.0$", "", regex=True)
    return df[df[FAMID_COL].isin(TARGET_FAMIDS)][[FAMID_COL] + COMPARE_COLS]

a = prep(FILE_A).rename(columns={c: f"{c}_A" for c in COMPARE_COLS})
b = prep(FILE_B).rename(columns={c: f"{c}_B" for c in COMPARE_COLS})

merged = pd.merge(a, b, on=FAMID_COL, how="outer", indicator=True)
for col in COMPARE_COLS:
    merged[f"{col}_match"] = merged[f"{col}_A"] == merged[f"{col}_B"]

with pd.ExcelWriter(
    f"{custom_path}/{filename}",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace",
) as writer:
    merged.to_excel(writer, sheet_name="IQ_comparison", index=False)