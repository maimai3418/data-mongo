# data-mongo

將心理衡鑑量表資料（xlsx）依量表類型拆分，驗證後匯入 MongoDB。

---

## 資料夾結構

```
data-mongo/
├── .env
├── .gitignore
├── requirements.txt
├── config.py
├── main.py              # 標準匯入（需 record_date）
├── age_main.py          # 年齡匯入（只需 age，自動估算 record_date）
├── delete_main.py       # 依 famid + record_date 刪除資料
└── src/
    ├── __init__.py
    ├── reader.py          # 讀取 xlsx（標準匯入用）
    ├── transformer.py     # 欄位驗證、規則比對、資料拆分
    ├── importer.py        # MongoDB upsert
    ├── age_matcher.py     # 年齡衝突檢查、record_date 估算
    ├── error_writer.py    # 錯誤紀錄寫入 errors.xlsx
    ├── no_date_writer.py  # 年齡衝突紀錄寫入 no_date_records.xlsx
    ├── import_logger.py   # 匯入統計寫入 import_log.xlsx（log + pivot）
    ├── logger.py          # console 摘要輸出
    └── utils/
        ├── select_collections.py  # 互動選擇量表
        └── wait_and_retry.py      # 檔案鎖定重試
```

---

## 使用步驟

### 1. 建立虛擬環境並安裝套件

```bash
python -m venv venv

# macOS/Linux
source venv/bin/activate

# Windows
venv\Scripts\activate

pip install -r requirements.txt
```

### 2. 設定 .env

```
MONGO_URI=mongodb://username:password@localhost:27017
MONGO_DB=your_db_name
```

### 3. 放入資料檔案

將 `.xlsx` 檔案放入專案根目錄，預設檔名為 `import_data.xlsx`。

Excel 檔案須包含名為 `import` 的 sheet，欄位需求依匯入模式不同：

| 模式 | 必要欄位 | 入口 |
|------|---------|------|
| 標準匯入 | `famid`、`record_date`、量表欄位 | `main.py` |
| 年齡匯入 | `famid`、`age`、量表欄位 | `age_main.py` |

### 4. 執行匯入

```bash
# 標準匯入（有 record_date）
python main.py

# 年齡匯入（只有 age，無 record_date）
python age_main.py

# 刪除資料
python delete_main.py
```

啟動後會互動選擇要處理的量表。

---

## 匯入邏輯

### 標準匯入（main.py）

1. 讀取 Excel `import` sheet，篩選有 `famid` + `record_date` 的行
2. 依 `config.py` 中的 `COLLECTION_MAP` 拆分欄位到各量表
3. 驗證每個欄位值（型別、範圍）
4. 量表欄位全為空值 / 999 的行自動跳過
5. 驗證失敗的行寫入 `errors.xlsx`
6. 通過驗證的資料 upsert 至 MongoDB

### 年齡匯入（age_main.py）

適用於只有 `age`（歲數）但沒有確切 `record_date` 的資料：

1. 讀取 Excel，篩選有 `famid` + `age` 的行
2. 驗證量表欄位值（同標準匯入）
3. 從 `Participants` collection 查詢 `birth_date`，計算已存在紀錄的年齡
4. 若已存在紀錄的年齡與匯入的 age 相差 ≤ 1 歲 → 視為潛在重複，寫入 `no_date_records.xlsx`
5. 無衝突的資料：以 `birth_date + age` 估算 `record_date`，標記 `est_record_date: 1`，upsert 至 MongoDB
6. 每筆匯入資料自動加上 `raw_{scale}_age` 欄位保留原始年齡

---

## 驗證規則

- 唯一鍵：`famid` + `record_date` + `role`（已存在不覆寫，使用 `$setOnInsert`）
- 欄位值為 `999`、空白、`N/A` 時存入 `null`，跳過範圍驗證
- 欄位值超出有效範圍 → 整筆不匯入，記錄至 `errors.xlsx`

### COLLECTION_MAP 設定格式

```python
# list 格式（簡單量表）
"GSQ": SHARED_FIELDS + ["gsq_"]

# dict 格式（需要額外欄位或個別規則的量表）
"ASRI": {
    "prefixes": ["asri_"],
    "extra_cols": { "欄位名": {"type": "float", "min": 0} },
    "field_rules": { "asri_48": {"type": "int", "range": (0, 1)} },
    "default_rule": {"type": "int", "range": (0, 3)},
}
```

---

## 輸出檔案

| 檔案 | 說明 |
|------|------|
| `errors.xlsx` | 驗證失敗的資料（每次執行新增 sheet） |
| `no_date_records.xlsx` | 年齡衝突或無法估算日期的資料（含 timestamp） |
| `import_log.xlsx` | **log** sheet：每次匯入的完整紀錄；**pivot** sheet：依量表分組的統計數量 |

---

## 支援量表

GSQ、SNAP4、CBCL、AQ、SRS、SAICA、SCQ、CAST、BRIEF、CPRS、CTRS、SDQ、PMI、MPBI、DPBI、APGAR、EPQ、YSR、SUB、ASRI、ASRS、CES-D、MPI
