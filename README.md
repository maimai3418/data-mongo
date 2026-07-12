# data-mongo

將心理衡鑑量表與神經心理測驗資料（xlsx）依工具類型拆分，驗證後匯入 MongoDB。除通用量表匯入管線（general/）外，另有 ADOS、ADI-R、CPT、CANTAB、IQ、KE、Participants 各自的專用管線，以及匯出、統計、重複檢查與 famid 比對等周邊工具。

---

## 資料夾結構

```
data-mongo/
├── .env
├── .gitignore
├── requirements.txt
├── config.py                # COLLECTION_MAP / VALID_RANGE 量表欄位與驗證規則
├── test_config.py           # 驗證 COLLECTION_MAP 與 VALID_RANGE 對應一致性
│
│  # === famid 比對 / 對照 ===
├── famid_compare.py         # 比對兩份 participants JSON 的 famid 差異（compare_famid.xlsx）
├── famid_project_map.py     # 掃描資料夾，建立 famid → 計畫對照（CSV）
├── reconcile_famid.py       # tool 資料夾 vs participants JSON 雙向比對（CSV）
│
├── general/                 # === 量表匯入 / 刪除 / 匯出 主程式 ===
│   ├── main.py              # 標準匯入（需 record_date，可選填計畫代碼）
│   ├── age_main.py          # 年齡匯入（只需 age，自動估算 record_date）
│   ├── delete_main.py       # 依 famid + record_date 刪除資料
│   ├── export_main.py       # 匯出 MongoDB 資料為 JSON
│   └── precheck_upload.py   # 上傳前重複檢查（唯讀，不寫入；輸出 precheck_report.xlsx）
│
├── participants/            # === Participants 主參與者清單 ===
│   ├── p_precheck.py        # 跨來源 birth_date / sex / group 衝突偵測
│   └── p_add_new.py         # 新 famid 送 staging；promote 模式搬入 Participants
│
├── ados/                    # === ADOS 管線 ===
│   ├── ados_config.py       # 設定：M1~M4 collection 對應、999→null、field.json 規則
│   ├── ados_field_rename.py # 匯入前欄位驗證與重命名（M1~M4 工作表）
│   ├── ados_import.py       # 依工作表匯入 ADOS_M1~M4（famid + record_date，不看 role）
│   ├── ados_delete.py       # 依 famid + record_date 刪除（支援 dry-run）
│   └── fields/              # 各 module 的 field.json
│
├── adir/                    # === ADI-R 管線 ===
│   ├── adir_config.py       # 設定：單一 ADIR collection、月齡欄位 999 為合法值
│   ├── adir_field_rename.py # 匯入前欄位驗證與重命名（Full / Current + version 標註）
│   ├── adir_import.py       # 匯入 ADIR
│   └── fields/              # field.json
│
├── cpt/                     # === CPT 管線 ===
│   ├── cpt_config.py        # 設定：field.json 規則（不做全域 999→null）
│   ├── cpt_convert.py       # 原始 txt → xlsx
│   ├── cpt_merge_raw.py     # 多工作表原始資料以欄位指紋去重合併
│   ├── cpt_backfill_dates.py# 以 merged 檔回填缺失 record_date（不碰 DB）
│   ├── cpt_precheck.py      # 匯入前彙整檢查（famid + age 去重，不碰 DB）
│   ├── cpt_import.py        # 三層去重複後匯入 CPT
│   └── fields/              # field.json
│
├── cantab/                  # === CANTAB 管線 ===
│   ├── cantab_config.py     # 設定：field.json 規則、年齡欄位別名
│   ├── cantab_rename.py     # 原始長欄名 → DB 代碼（cantab_new_fields.xlsx 對照）
│   ├── cantab_precheck.py   # 匯入前彙整檢查（famid + age 去重，不碰 DB）
│   ├── cantab_backfill.py   # session_start_time 回填嘗試（僅報表，不碰 DB）
│   ├── cantab_import.py     # 三層去重複後匯入 CANTAB
│   └── fields/              # field.json + cantab_new_fields.xlsx
│
├── iq/                      # === IQ 管線（匯入腳本尚未實作）===
│   └── iq_precheck.py       # 匯入前彙整檢查（famid + age 去重，資料驅動模式）
│
├── ke/                      # === KE (K-SADS-E) 管線 ===
│   ├── ke_backup.py         # 全部 KE 子 collection 備份為 Extended JSON
│   ├── ke_import.py         # 備份 JSON 還原上傳（預設 dry-run）
│   ├── ke_delete.py         # 依 collection + famid + record_date 刪除（預設 dry-run）
│   ├── ke_key_check.py      # 全部 KE 子 collection 的 unique key 差異矩陣
│   ├── ke_test.py           # 比對兩個 collection 的 unique key
│   └── fields/              # *_ke_fields_grouped.json（collection 白名單）
│
├── stat/                    # === DB 統計 / 查詢 ===
│   ├── collection_stat.py   # 依 collection 統計 famid 數量、格式分類與重複分布
│   └── check_stat.py        # 互動查詢單一 collection 的 document 數與 role 分組
│
├── tools/                   # === 通用 Excel / 檔案工具 ===
│   ├── scan_column.py       # 掃描資料夾，找出含特定欄位關鍵字的檔案（matched_files.xlsx）
│   ├── cell_extract.py      # 讀資料夾內所有 xls 的指定儲存格，整理成新 xlsx（cell_extract.xlsx）
│   ├── sheet_compare.py     # 比對同一 xlsx 內兩個工作表的欄位與逐列差異（comparison.xlsx）
│   ├── compare_columns.py   # 比對多個 xlsx 的標題列，輸出欄位×來源矩陣
│   └── check_duplicate_keys.py # COLLECTION_MAP 各 collection 的 composite key 重複檢查（唯讀）
│
└── src/                     # === 共用模組（寫新腳本先重用，不要重寫）===
    ├── __init__.py
    ├── reader.py                  # 讀取 xlsx（標準匯入用）
    ├── transformer.py             # 欄位驗證、規則比對、資料拆分
    ├── importer.py                # MongoDB 連線（get_db）與 upsert 分類
    ├── age_matcher.py             # 年齡衝突檢查、record_date 估算
    ├── conflict_writer.py         # value_conflict 差異寫入 conflicts.xlsx
    ├── error_writer.py            # 錯誤紀錄寫入 errors.xlsx
    ├── no_date_writer.py          # 無日期 / 年齡衝突紀錄寫入 no_date_records.xlsx
    ├── import_logger.py           # 匯入統計寫入 import_log.xlsx（log + pivot）
    ├── logger.py                  # console 摘要輸出
    └── utils/
        ├── select_collections.py  # 互動選擇量表
        ├── input_project_code.py  # 互動輸入計畫代碼（選填）
        └── wait_and_retry.py      # 檔案鎖定重試（所有寫 xlsx 的腳本統一使用）
```

**共通規範**（所有腳本一致）：

- DB 連線一律走 `src/importer.py` 的 `get_db()`（讀 `.env` 的 `MONGO_URI` / `MONGO_DB`），不各自建 `MongoClient`
- 從子資料夾執行的腳本，開頭都有 sys.path bootstrap，一律**從專案根目錄執行**
- 所有寫 xlsx 的動作都包 `src/utils/wait_and_retry.py`：目標檔案被 Excel 開啟時提示關閉後重試，不會直接 crash
- 錯誤 / 衝突報表以 openpyxl 格式化輸出，命名格式 `{YYYYMMDD}_{工具}_{用途}.xlsx`

---

## 安裝與設定

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

將 `.xlsx` 檔案放入專案根目錄，匯入預設檔名為 `import_data.xlsx`。

Excel 檔案須包含名為 `import` 的 sheet，欄位需求依匯入模式不同：

| 模式     | 必要欄位                                 | 入口                  |
| -------- | ---------------------------------------- | --------------------- |
| 標準匯入 | `famid`、`record_date`、`role`、量表欄位 | `general/main.py`     |
| 年齡匯入 | `famid`、`age`、量表欄位                 | `general/age_main.py` |

---

## 功能總覽

### 量表匯入 / 匯出 / 統計

| 指令                                | 用途                                       | 輸出                              |
| ----------------------------------- | ------------------------------------------ | --------------------------------- |
| `python general/main.py`            | 標準匯入（有 record_date）                 | MongoDB、`errors.xlsx`、`conflicts.xlsx`、`import_log.xlsx` |
| `python general/age_main.py`        | 年齡匯入（只有 age）                       | MongoDB、`no_date_records.xlsx`   |
| `python general/delete_main.py`     | 依 famid + record_date 刪除資料            | MongoDB                           |
| `python general/export_main.py`     | 匯出量表資料為 JSON                        | `exports/YYYYMMDD/*.json`         |
| `python general/precheck_upload.py [檔案]` | 上傳前重複檢查（唯讀）             | `precheck_report.xlsx`            |
| `python stat/collection_stat.py`    | collection famid 統計                      | console 或 `collection_stat.xlsx` |
| `python stat/check_stat.py`         | 查單一 collection 的數量與 role 分組       | console                           |
| `python test_config.py`             | 驗證 config 設定一致性                     | console（PASS/FAIL）              |
| `python tools/check_duplicate_keys.py` | COLLECTION_MAP 各 collection 的 key 重複檢查（唯讀） | `{TODAY}_dup_key_report.xlsx` |

### 各工具專用管線

| 指令                                    | 用途                                         | 輸出                                  |
| --------------------------------------- | -------------------------------------------- | ------------------------------------- |
| `python participants/p_precheck.py <dir>` | Participants 跨來源欄位衝突偵測            | `{TODAY}_p_precheck_conflict.xlsx`    |
| `python participants/p_add_new.py <dir>` | 新 famid 送 staging / promote               | `{TODAY}_p_new_candidates.xlsx`、MongoDB |
| `python ados/ados_field_rename.py <dir>` | ADOS 匯入前欄位驗證與重命名                 | 改名後 xlsx、`{TODAY}_ados_error.xlsx` |
| `python ados/ados_import.py <dir>`       | ADOS 匯入（M1~M4 → ADOS_M1~M4）             | MongoDB、`{TODAY}_ados_import_error.xlsx` |
| `python ados/ados_delete.py`             | ADOS 刪除（讀 ados_delete.xlsx）            | MongoDB                               |
| `python adir/adir_field_rename.py <dir>` | ADI-R 匯入前欄位驗證與重命名                | 改名後 xlsx、`{TODAY}_adir_rename_error.xlsx` |
| `python adir/adir_import.py <dir>`       | ADI-R 匯入（單一 ADIR collection）          | MongoDB、`{TODAY}_adir_import_error.xlsx` |
| `python cpt/cpt_convert.py [txt]`        | CPT 原始 txt 解析成 xlsx                    | `cpt_parsed.xlsx`                     |
| `python cpt/cpt_merge_raw.py`            | CPT 原始多工作表指紋去重合併                | `*_CCPT_merged.xlsx`                  |
| `python cpt/cpt_backfill_dates.py`       | CPT 缺失 record_date 回填（不碰 DB）        | `*_backfilled.xlsx`                   |
| `python cpt/cpt_precheck.py <dir>`       | CPT 匯入前彙整檢查（不碰 DB）               | `cpt_precheck.xlsx`                   |
| `python cpt/cpt_import.py <dir>`         | CPT 三層去重複匯入                          | MongoDB、`{TODAY}_CPT_import_error.xlsx`、`{TODAY}_CPT_conflict_report.xlsx`、`no_date_records.xlsx` |
| `python cantab/cantab_rename.py <dir>`   | CANTAB 原始長欄名 → DB 代碼                 | `output/cantab_renamed/*.xlsx`        |
| `python cantab/cantab_precheck.py <dir>` | CANTAB 匯入前彙整檢查（不碰 DB）            | `cantab_precheck.xlsx`                |
| `python cantab/cantab_backfill.py`       | CANTAB session_start_time 回填嘗試（僅報表）| `{TODAY}_CANTAB_sst_backfill.xlsx`    |
| `python cantab/cantab_import.py <dir>`   | CANTAB 三層去重複匯入                       | MongoDB、`{TODAY}_CANTAB_import_error.xlsx`、`{TODAY}_CANTAB_conflict_report.xlsx`、`no_date_records.xlsx` |
| `python iq/iq_precheck.py <dir>`         | IQ 匯入前彙整檢查（不碰 DB）                | `iq_precheck.xlsx`                    |
| `python ke/ke_backup.py`                 | KE 全部子 collection 備份                   | `ke/backup/{YYYYMMDD}/*.json`         |
| `python ke/ke_import.py`                 | KE 備份還原上傳（預設 dry-run）             | MongoDB                               |
| `python ke/ke_delete.py`                 | KE 刪除（讀 ke_delete.xlsx，預設 dry-run）  | MongoDB                               |
| `python ke/ke_key_check.py`              | KE 全部子 collection 的 key 差異矩陣        | `{TODAY}_ke_key_check.xlsx`           |
| `python ke/ke_test.py`                   | 比對兩個 collection 的 unique key           | `{TODAY}_ke_compare.xlsx`             |

### famid 比對 / 通用工具

| 指令                                | 用途                                       | 輸出                              |
| ----------------------------------- | ------------------------------------------ | --------------------------------- |
| `python famid_compare.py`           | 比對兩份 participants JSON                 | `compare_famid.xlsx`              |
| `python famid_project_map.py <dir>` | famid → 計畫對照                           | `famid_project_map.csv` 等        |
| `python reconcile_famid.py <dir>`   | tool 資料夾 vs participants JSON 比對      | `orphan_*.csv`、`inactive_*.csv`  |
| `python tools/scan_column.py <dir>` | 找出含特定欄位關鍵字的檔案                 | `matched_files.xlsx`              |
| `python tools/cell_extract.py <dir>`| 讀資料夾內 xls 的指定儲存格整理成表        | `cell_extract.xlsx`               |
| `python tools/sheet_compare.py <檔案>` | 比對同一 xlsx 內兩個工作表              | `comparison.xlsx`                 |
| `python tools/compare_columns.py <dir>` | 比對多個 xlsx 的標題列                | 指定的報表 xlsx                   |

所有指令一律從專案根目錄執行。量表匯入、刪除、匯出、檢查類指令啟動後會互動選擇要處理的量表；各工具管線指令改用命令列參數（多數支援 `--dry-run` 或預設 dry-run）。

---

## 匯入 / 刪除 / 匯出

### 標準匯入（general/main.py）

1. 輸入計畫代碼（選填，按 Enter 跳過）
2. 互動選擇要上傳的量表
3. 讀取 Excel `import` sheet，過濾完全空白的列
4. 依 `config.py` 中的 `COLLECTION_MAP` 拆分欄位到各量表
5. 驗證每個欄位值（型別、範圍）
6. 量表欄位全為空值 / 999 的行自動跳過
7. 必要欄位（`famid` / `record_date` / `role`）檢查，缺少則報錯（見下方說明）
8. 驗證失敗的行寫入 `errors.xlsx`
9. 通過驗證的資料依 unique key（`role` + `famid` + `record_date`）與 DB 現有資料比對分類：
   - **new_insert**：DB 無對應 key → 插入；若有輸入計畫代碼，每筆加上 `research_project_code` 欄位
   - **exact_dup**：key 相同且所有欄位值相同 → 跳過（計數）
   - **value_conflict**：key 相同但欄位值不同 → **不覆蓋 DB**，欄位層級差異（new_value vs db_value 並列）匯出至 `conflicts.xlsx`

#### 必要欄位檢查（famid / record_date / role）

分兩種狀況處理，缺少時顯示對應錯誤訊息：

1. **整欄不存在**（Excel 沒有該欄）→ 讀檔時即中止匯入，印出 `匯入中止：Excel 缺少必要欄位欄 {欄位名}`
2. **某列缺值**（有量表資料但缺其中一欄）→ 該列寫入 `errors.xlsx`，錯誤訊息為 `缺少必要欄位: {欄位名}`

> `record_date` 無法解析的值視同缺值，由上述「某列缺值」邏輯報錯；完全空白的列則直接略過，不視為錯誤。

### 年齡匯入（general/age_main.py）

適用於只有 `age`（歲數）但沒有確切 `record_date` 的資料：

1. 讀取 Excel，篩選有 `famid` + `age` 的行
2. 驗證量表欄位值（同標準匯入）
3. 從 `Participants` collection 查詢 `birth_date`，計算已存在紀錄的年齡
4. 若已存在紀錄的年齡與匯入的 age 相差 ≤ 1 歲 → 視為潛在重複，寫入 `no_date_records.xlsx`
5. 無衝突的資料：以 `birth_date + age` 估算 `record_date`，標記 `est_record_date: 1`，upsert 至 MongoDB
6. 每筆匯入資料自動加上 `raw_{scale}_age` 欄位保留原始年齡

### 刪除資料（general/delete_main.py）

依 `famid` + `record_date` 從選定的量表 collection 刪除資料。

### 匯出 JSON（general/export_main.py）

將 MongoDB 中的量表資料匯出為 JSON 檔案：

1. 互動選擇要匯出的量表（或全部）
2. 自動建立 `exports/YYYYMMDD/` 資料夾（以當日日期命名）
3. 每個量表匯出為獨立的 `{collection_name}.json`

---

## 統計 / 檢查

### Collection 統計（stat/collection_stat.py）

互動式輸入 collection 名稱，輸出以下統計：

1. **distinct famid 數量** — 該 collection 內不重複的 famid 總數
2. **famid 格式分類（distinct）** — 分別計數
   - 2~5 位數字、開頭為 `1`
   - 5 位數字、開頭為 `3 / 5 / 6 / 8 / 9`
   - 6 位數字、開頭為 `4`
   - 不符合上述規則者歸入「其他」
3. **famid 結尾數字分布（distinct）** — 各 distinct famid 依末位數字 `0~9` 計數（獨立統計，與格式分類分開）
4. **重複出現次數分布** — 各 famid 在此 collection 中出現的次數分布（例：`count=1: 3`、`count=2: 4`、`count=3: 1` …）

輸入 `all` 或 `export` → 依 `COLLECTION_MAP` 掃描所有 collection，輸出至 `exports/collection_stat/{MMDD_HHMMSS}_collection_stat.xlsx`：

- **summary** sheet：每列一個 collection，欄位含 `total_records`、`distinct_famid`、各格式分類計數、`其他`
- **by_last_digit** sheet：每列一個 collection，欄位為 `distinct_famid` 與 `0 結尾`~`9 結尾`
- **duplicates** sheet：每列一個 collection，欄位為 `count=1`、`count=2`、…（依實際出現的最大次數動態展開）

### 單一 Collection 查詢（stat/check_stat.py）

互動式輸入 collection 名稱，顯示總 document 數量，並依 `role` 分組計數。輸入 `q` 離開。

### 上傳前重複檢查（general/precheck_upload.py）

在**不寫入 MongoDB** 的前提下，拿一份新的 Excel 資料跟資料庫現有資料比對，用來 double-check「會不會重複上傳」以及「日期、計畫是否相符」。完全沿用 `main.py` 的流程（輸入計畫代碼 → 選擇量表 → `read_xlsx` → `split_by_collection`），所以檢查的就是「真的會被上傳」的那批資料，且不更動任何現有模組。

對資料庫**只有查詢（find）、沒有任何寫入**，唯一的輸出是本機報表 `precheck_report.xlsx`。

```bash
python general/precheck_upload.py              # 預設讀 import_data.xlsx
python general/precheck_upload.py 其他資料.xlsx  # 指定要檢查的檔案
```

依唯一鍵 `role` + `famid` + `record_date` 逐筆比對，產出 5 個工作表：

1. **Summary** — 各量表：檢查筆數 / 新資料(會上傳) / 已存在(會略過) / 計畫不符
2. **Duplicates** — 已存在的紀錄，含 DB 計畫 vs 本次計畫與是否相符（不符標黃）
3. **Project_Mismatch** — 只列計畫不符的紀錄（因 upsert 用 `$setOnInsert`，重傳不會更新既有計畫，需特別注意）
4. **New_Rows** — 會被新增上傳的紀錄
5. **Date_Check** — 每個 famid：本檔 record_date vs DB 既有 record_date vs 本檔多出的日期

### Config 一致性測試（test_config.py）

驗證 `config.py` 設定是否一致，可用 `python test_config.py` 直接執行（印出 PASS/FAIL）或以 `pytest` 執行：

- 每個 `COLLECTION_MAP` 的 key 都能解析到驗證規則（dict 自帶 `default_rule` 或存在於 `VALID_RANGE`）
- `VALID_RANGE` 沒有多餘、不在 `COLLECTION_MAP` 的 key
- `COLLECTION_MAP` 的值格式正確（list 或含 `prefixes` 的 dict）
- `VALID_RANGE` 每個規則都有 `type`，`type=int` 時需有合理的 `range`

---

## 各工具專用管線

以下工具不走 `config.py` 的 COLLECTION_MAP，各自有專用 config 與 `fields/` 下的 field.json（型別 / 範圍規則）。共通慣例：

- 匯入的 unique key 為 `famid` + `record_date`（**不看 role**，與量表管線不同），已存在不覆寫（`$setOnInsert`）
- item 欄位全為空的列略過不上傳
- `--dry-run` 只驗證不寫入；`--project-code` 可在新增筆數上標註計畫代碼

### Participants（participants/）

- **p_precheck.py**：偵測同一 famid 跨來源的 `birth_date` / `sex` / `group` 差異。每個欄位獨立判定 base（DB 已有值以 DB 為準，否則取第一筆有值的來源），衝突者掛 `diff_*` 與 `*_conflict` 旗標，匯出 `{TODAY}_p_precheck_conflict.xlsx`。export / write 模式皆需 DB 連線。
- **p_add_new.py**：將尚未存在於 `Participants` 的 famid 送入 staging collection；staging 內已存在者比對差異並標記，完全相同則跳過。`--mode promote` 把 staging 中已補 `group` 者搬入 `Participants`。

### ADOS（ados/）

兩階段：先 rename、後 import。999 → null，其餘值原樣上傳。

1. **ados_field_rename.py**：驗證各檔案的 M1~M4 工作表是否含該 module 完整項目欄位（依 `ados/fields/` field.json），檢查 famid / record_date（含 fallback 組合邏輯），重命名欄位為標準名。預設覆蓋原檔，`--no-overwrite` 另存 `_renamed.xlsx`。
2. **ados_import.py**：依工作表名稱匯入對應 collection（M1 → `ADOS_M1` …）。`--modules M1 M3` 指定 module。
3. **ados_delete.py**：讀 `ados_delete.xlsx`（famid、record_date、選填 module 欄），從指定 module collection 刪除；`--dry-run` 預覽、`-y` 略過確認。

### ADI-R（adir/）

與 ADOS 同款兩階段，差異：單一 `ADIR` collection 不分 module；**月齡欄位的 999 為合法值**（field.json range 涵蓋 999 者先照 range 驗證，不轉 null）；rename 階段依工作表名稱（`ADIR_SHEET_VERSION_MAP`）自動標註 `version`（full / current）。

1. **adir_field_rename.py**：famid 缺少時嘗試從 family+id 或 site+fam+id 組合；record_date 缺少時嘗試從 int_y/int_m/int_d 組合；只保留 field.json 定義的欄位並依其排序。
2. **adir_import.py**：匯入單一 `ADIR` collection。

### CPT（cpt/）

原始資料整理 → precheck → 匯入：

1. **cpt_convert.py**：把 CPT 匯出的原始 txt（每筆固定 8 行區塊）解析成 xlsx，自動偵測有無 `name` 欄、famid 去除 `.0` 殘留。
2. **cpt_merge_raw.py**：多張原始工作表合併去重。famid 不完全可靠，改以分數欄位指紋（record_date + 39 個分數欄）判定同一筆施測；famid 不一致的組另開 `famid_conflicts` 工作表供人工核對。
3. **cpt_backfill_dates.py**：以 merged 檔比對回填 cleaned 檔缺失的 record_date（tier1：39 欄全等；tier2：排除重算欄後 29 欄全等），已有日期不覆寫、日期不合者出 `date_mismatch` 報表。不碰 DB。
4. **cpt_precheck.py**：多檔彙整，以 famid + age 為 key 去重（欄名一律轉小寫比對），輸出 Summary / Files / Duplicates / Conflicts / Records 五工作表。
5. **cpt_import.py**：三層去重複後匯入 `CPT`（見下方「三層去重複」）。

### CANTAB（cantab/）

1. **cantab_rename.py**：依 `cantab/fields/cantab_new_fields.xlsx` 對照表把原始長欄名（如 "MOT Mean latency"）改成 DB 代碼（如 "MOTmL"）；認不得的欄位保留原名並在報告標黃（非預期值必須明確浮現）。傳入資料夾時彙整成單一 xlsx。
2. **cantab_precheck.py**：同 CPT precheck（famid + age 去重），比對前先把各檔欄位對齊 field.json（缺欄 / 空值以 999 填入，僅供比對用）。
3. **cantab_backfill.py**：對缺 `session_start_time` 的列，以同 famid 的錨點日期（datapool 內 + DB 既有紀錄）配合 `Participants` birth_date + 年齡驗證後聚類，提議回填日期或匯出候選供人工判定。僅輸出報表，不改原始檔、不寫 DB。
4. **cantab_import.py**：三層去重複後匯入 `CANTAB`。無 record_date 但有年齡欄位者：`record_date=null` + `age_suspect=true`，年齡存入 `raw_cantab_age`。

**三層去重複**（CPT / CANTAB 匯入共用，CLAUDE.md §5）：

1. **hard dedup**：同批次內 famid + record_date + 所有量表欄位值完全相同 → 只留一筆
2. **retest 保留**：famid 相同但 record_date 為 DB 中新日期 → 視為重測直接新增
3. **outer join 人工審查**：key 與 DB 相同但欄位值不同 → 兩邊皆保留，匯出 `{TODAY}_{工具}_conflict_report.xlsx`

無 record_date 的列不套用三層比對：以 famid + 欄位值簽章做 hard dedup，其餘一律新增（`record_date=null`、`flag_no_date=true`），全部記錄到 `no_date_records.xlsx` 供日期回填管線參考。CPT / CANTAB **不做全域 999→null 轉換**（999 非通用 sentinel）。

### IQ（iq/）

- **iq_precheck.py**：同 CPT/CANTAB precheck。IQ 目前沒有 field.json，採**資料驅動**模式（比對欄位 = 讀到的所有欄位，扣掉 famid/age/sex/birth_date/record_date）；之後補上 `iq/fields/*iq*fields.json` 會自動改用。
- 匯入腳本（iq_import.py / iq_rename.py）尚未實作。

### KE（ke/）

KE 拆成多個子 collection（清單來自 `ke/fields/*_ke_fields_grouped.json`，取檔名最新一份作白名單）。刪除 / 上傳類預設 dry-run，需 `--execute` + 確認才動 DB：

- **ke_backup.py**：全部 KE 子 collection 以 bson Extended JSON 完整下載到 `ke/backup/{YYYYMMDD}/`，`_id` 與日期無損保存，空 collection 也輸出空陣列。
- **ke_import.py**：把備份 JSON 上傳回對應 collection（檔名 = collection 名，需在白名單內）。以 `_id` 做 ReplaceOne upsert，重複執行冪等。
- **ke_delete.py**：讀 `ke/ke_delete.xlsx` 的 `delete` 工作表（collection / famid / record_date），僅允許白名單內 collection。
- **ke_key_check.py**：比對所有子 collection 的 unique key（famid + record_date）差異，輸出 summary + diff_matrix。
- **ke_test.py**：比對 CONFIG 中指定的兩個 collection 的 unique key（summary / A_only / B_only）。

---

## famid 比對 / 對照

### 兩份 JSON 比對（famid_compare.py）

比對同資料夾下 `db.json`（A，正式/基準）與 `test.json`（B，測試）兩份 participants 的 famid 重複與差異狀況，輸出 `compare_famid.xlsx`，含 5 個工作表：Summary、Duplicates（各檔重複 famid）、Both_Diff（共同 famid 逐欄差異）、Only_in_A、Only_in_B。

### famid → 計畫對照（famid_project_map.py）

以「最外層資料夾 = 一個計畫」為前提，遞迴掃描資料夾下所有 `.xlsx` / `.xls`（每檔只取第一個 sheet）的 `famid` 欄位，統計每個 famid 出現在哪些計畫、跨幾個計畫。

```bash
python famid_project_map.py <tool_folder>   # 指到「各計畫資料夾的上一層」
```

輸出：
- `famid_project_map.csv` — 每列一個 famid：`n_projects` / `projects` / `n_files` / `sources`
- `project_summary.csv` — 每個計畫的 famid 數、檔案數

### tool vs participants 雙向比對（reconcile_famid.py）

比對 tool 資料夾（遞迴所有 `.xlsx` / `.xls`，每檔第一個 sheet）的 famid 聯集與一份 participants JSON 的 famid。participants JSON 結構自動偵測（陣列、JSONL/NDJSON，或包在 `data` / `participants` / `results` / `items` / `records` key 內）。

```bash
python reconcile_famid.py <tool_folder> [participants.json]
# 第二參數省略時，預設讀執行目錄下的 participants.json
```

輸出：
- `orphan_famid_report.csv` — tool 有、participants 無
- `inactive_participant_report.csv` — participants 有、tool 都無

---

## 通用工具

### 欄位關鍵字掃描（tools/scan_column.py）

遞迴掃描資料夾下的 `.xlsx` / `.xls` / `.xlsm` / `.csv` / `.tsv`，找出欄名含特定關鍵字（預設 `time` / `test` / `study`）的檔案與 sheet，輸出 `matched_files.xlsx`。

```bash
python tools/scan_column.py <dir>
python tools/scan_column.py <dir> -o out.xlsx --keywords date time --ignore .git venv
```

### 指定儲存格擷取（tools/cell_extract.py）

掃描資料夾下所有 `.xls`（也支援 `.xlsx` / `.xlsm`），依「儲存格位址 → 輸出欄名」對應表，讀出每個檔案的特定儲存格，整理成一份新的 xlsx（每個檔案一列，第一欄附 `source_file` 來源檔名）。對應表由使用者自行定義：用 `--map A1=id B5=name`（行內，可多組）或 `--map-file`（JSON）指定。

```bash
python tools/cell_extract.py <dir> --map A1=id B5=name C2=score   # 行內定義對應
python tools/cell_extract.py <dir> --map-file mapping.json        # JSON：{"A1": "id", "B5": "name"}
python tools/cell_extract.py <dir> --map A1=id --sheet Data       # 指定工作表（名稱或索引，預設第一個）
python tools/cell_extract.py <dir> --map A1=id -r                 # 遞迴子資料夾（預設只掃最外層）
python tools/cell_extract.py <dir> --map A1=id -o out.xlsx
```

- 超出範圍的儲存格留空、不報錯；整數型數值去除 `.0`、字串去頭尾空白
- 讀檔失敗的檔案仍保留該列，於 `_note` 欄記錄原因並標色

### 工作表比對（tools/sheet_compare.py）

比對同一個 xlsx 內兩個工作表（A、B）的欄位與逐列數值差異，輸出 `comparison.xlsx`（Summary / Columns / Cell_Diff / Only_in_A / Only_in_B 五工作表）。

```bash
python tools/sheet_compare.py data.xlsx                  # 預設取前兩個工作表，依位置對齊
python tools/sheet_compare.py data.xlsx --sheets A B     # 指定工作表名稱
python tools/sheet_compare.py data.xlsx --key famid role # 依（複合）鍵對齊
```

### 標題列比對（tools/compare_columns.py）

比對多個 xlsx 的標題列，輸出欄位 × 來源檔案的對照矩陣，快速看出各來源檔的欄位差異。

### Composite key 重複檢查（tools/check_duplicate_keys.py）

以 `config.SHARED_FIELDS`（role + famid + record_date）為 key，掃描 DB 中 `COLLECTION_MAP` 全部 collection，找出 key 重複的紀錄。只讀不改；有重複時輸出 `{TODAY}_dup_key_report.xlsx` 供人工判定。

```bash
python tools/check_duplicate_keys.py
```

---

## 驗證規則（量表管線）

- 唯一鍵：`famid` + `record_date` + `role`（已存在不覆寫，使用 `$setOnInsert`）
- 欄位值為 `999`、空白、`N/A` 時存入 `null`，跳過範圍驗證
- 欄位值超出有效範圍 → 整筆不匯入，記錄至 `errors.xlsx`

> 以上 999→null 與 role 規則**僅限量表管線**。各工具專用管線的規則見各自章節：ADOS 只有 999 轉 null、ADI-R 月齡欄位 999 為合法值、CPT / CANTAB 完全不做全域 999→null 轉換；unique key 皆不含 role。

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

所有寫 xlsx 的動作都經 `wait_and_retry`：目標檔案被 Excel 開啟時，會提示關閉後按 Enter 重試。

| 檔案                          | 說明                                                                                                               |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `errors.xlsx`                 | 量表管線驗證失敗的資料（每次執行新增 sheet）                                                                       |
| `conflicts.xlsx`              | 量表管線 value_conflict 欄位層級差異（new_value vs db_value 並列）                                                 |
| `no_date_records.xlsx`        | 年齡衝突、無日期或無法估算日期的資料（含 timestamp）                                                               |
| `import_log.xlsx`             | **log** sheet：每次匯入的完整紀錄（含 success / insert / skipped_dup 欄位）；**pivot** sheet：依量表分組的統計數量 |
| `exports/YYYYMMDD/*.json`     | 從 MongoDB 匯出的量表 JSON 檔案                                                                                    |
| `exports/collection_stat/*.xlsx` | collection famid 統計（summary / by_last_digit / duplicates）                                                  |
| `precheck_report.xlsx`        | 上傳前重複檢查報表（唯讀產出）：Summary / Duplicates / Project_Mismatch / New_Rows / Date_Check                    |
| `{TODAY}_dup_key_report.xlsx` | COLLECTION_MAP 各 collection 的 composite key 重複紀錄（output/ 下）                                               |
| `{TODAY}_ados_error.xlsx` / `{TODAY}_ados_import_error.xlsx` | ADOS rename / 匯入錯誤                                                    |
| `{TODAY}_adir_rename_error.xlsx` / `{TODAY}_adir_import_error.xlsx` | ADI-R rename / 匯入錯誤                                             |
| `cpt_precheck.xlsx` / `cantab_precheck.xlsx` / `iq_precheck.xlsx` | precheck 報表（Summary / Files / Duplicates / Conflicts / Records）  |
| `{TODAY}_CPT_import_error.xlsx` / `{TODAY}_CANTAB_import_error.xlsx` | CPT / CANTAB 匯入驗證錯誤                                          |
| `{TODAY}_CPT_conflict_report.xlsx` / `{TODAY}_CANTAB_conflict_report.xlsx` | 三層去重複第三層的衝突報告（BatchConflicts / ValueConflicts） |
| `{TODAY}_CANTAB_sst_backfill.xlsx` | CANTAB session_start_time 回填提議報表                                                                        |
| `cpt_parsed.xlsx`             | CPT 原始 txt 解析結果（單一 `CPT` sheet）                                                                          |
| `ke/backup/{YYYYMMDD}/*.json` | KE 子 collection 備份（Extended JSON）                                                                             |
| `{TODAY}_ke_key_check.xlsx` / `{TODAY}_ke_compare.xlsx` | KE key 差異矩陣 / 兩 collection key 比對                                                 |
| `{TODAY}_p_precheck_conflict.xlsx` / `{TODAY}_p_new_candidates.xlsx` | Participants 衝突偵測 / 新增候選報表                               |
| `compare_famid.xlsx`          | 兩份 participants JSON 比對結果                                                                                    |
| `famid_project_map.csv` / `project_summary.csv` | famid → 計畫對照與計畫摘要                                                              |
| `orphan_famid_report.csv` / `inactive_participant_report.csv` | tool vs participants 雙向比對結果                                        |
| `matched_files.xlsx`          | 含特定欄位關鍵字的檔案清單                                                                                         |
| `cell_extract.xlsx`           | 資料夾內各 xls 指定儲存格的擷取結果（每檔一列）                                                                    |
| `comparison.xlsx`             | 同一 xlsx 內兩工作表的比對結果                                                                                     |

---

## 支援量表

GSQ、SNAP4、CBCL、AQ、SRS、SAICA、SCQ、CAST、BRIEF、CPRS、CTRS、SDQ、PMI、MPBI、DPBI、APGAR、EPQ、YSR、SUB、ASRI、ASRS、CES-D、MPI、RBS-R、SSP、WFIRS-P、WFIRS-S、WHOQOL、CEQ、C-SBEQ、ERQ-CA、BRIEF-S、BRIEF-A、ERQ-A、RAADS-R、TAS-20、AAQOL、ARI、ESQ
