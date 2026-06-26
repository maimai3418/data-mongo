# data-mongo

將心理衡鑑量表資料（xlsx）依量表類型拆分，驗證後匯入 MongoDB，並提供匯出、統計、重複檢查與 famid 比對等周邊工具。

---

## 資料夾結構

```
data-mongo/
├── .env
├── .gitignore
├── requirements.txt
├── config.py                # COLLECTION_MAP / VALID_RANGE 量表欄位與驗證規則
│
│  # === 匯入 / 刪除 / 匯出 ===
├── main.py                  # 標準匯入（需 record_date，可選填計畫代碼）
├── age_main.py              # 年齡匯入（只需 age，自動估算 record_date）
├── delete_main.py           # 依 famid + record_date 刪除資料
├── export_main.py           # 匯出 MongoDB 資料為 JSON
│
│  # === 統計 / 檢查 ===
├── collection_stat.py       # 依 collection 統計 famid 數量、格式分類與重複分布
├── check_stat.py            # 互動查詢單一 collection 的 document 數與 role 分組
├── precheck_upload.py       # 上傳前重複檢查（唯讀，不寫入；輸出 precheck_report.xlsx）
├── test_config.py           # 驗證 COLLECTION_MAP 與 VALID_RANGE 對應一致性
│
│  # === famid 比對 / 對照 ===
├── famid_compare.py         # 比對兩份 participants JSON 的 famid 差異（compare_famid.xlsx）
├── famid_project_map.py     # 掃描資料夾，建立 famid → 計畫對照（CSV）
├── reconcile_famid.py       # tool 資料夾 vs participants JSON 雙向比對（CSV）
│
│  # === 通用工具 ===
├── scan_column.py           # 掃描資料夾，找出含特定欄位關鍵字的檔案（matched_files.xlsx）
├── cell_extract.py          # 讀資料夾內所有 xls 的指定儲存格，整理成新 xlsx（cell_extract.xlsx）
├── cpt_convert.py           # 把 CPT 匯出的原始 txt（每筆 8 行）解析成 xlsx（cpt_parsed.xlsx）
├── tree.py                  # 匯出資料夾結構為 Excel（directory_tree.xlsx）
│
└── src/
    ├── __init__.py
    ├── reader.py                  # 讀取 xlsx（標準匯入用）
    ├── transformer.py             # 欄位驗證、規則比對、資料拆分
    ├── importer.py                # MongoDB 連線與 upsert
    ├── age_matcher.py             # 年齡衝突檢查、record_date 估算
    ├── error_writer.py            # 錯誤紀錄寫入 errors.xlsx
    ├── no_date_writer.py          # 年齡衝突紀錄寫入 no_date_records.xlsx
    ├── import_logger.py           # 匯入統計寫入 import_log.xlsx（log + pivot）
    ├── logger.py                  # console 摘要輸出
    └── utils/
        ├── select_collections.py  # 互動選擇量表
        ├── input_project_code.py  # 互動輸入計畫代碼（選填）
        └── wait_and_retry.py      # 檔案鎖定重試
```

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

| 模式     | 必要欄位                                 | 入口          |
| -------- | ---------------------------------------- | ------------- |
| 標準匯入 | `famid`、`record_date`、`role`、量表欄位 | `main.py`     |
| 年齡匯入 | `famid`、`age`、量表欄位                 | `age_main.py` |

---

## 功能總覽

| 指令                                | 用途                                       | 輸出                              |
| ----------------------------------- | ------------------------------------------ | --------------------------------- |
| `python main.py`                    | 標準匯入（有 record_date）                 | MongoDB、`errors.xlsx`、`import_log.xlsx` |
| `python age_main.py`                | 年齡匯入（只有 age）                       | MongoDB、`no_date_records.xlsx`   |
| `python delete_main.py`             | 依 famid + record_date 刪除資料            | MongoDB                           |
| `python export_main.py`             | 匯出量表資料為 JSON                        | `exports/YYYYMMDD/*.json`         |
| `python collection_stat.py`         | collection famid 統計                      | console 或 `collection_stat.xlsx` |
| `python check_stat.py`              | 查單一 collection 的數量與 role 分組       | console                           |
| `python precheck_upload.py [檔案]`  | 上傳前重複檢查（唯讀）                     | `precheck_report.xlsx`            |
| `python test_config.py`             | 驗證 config 設定一致性                     | console（PASS/FAIL）              |
| `python famid_compare.py`           | 比對兩份 participants JSON                 | `compare_famid.xlsx`              |
| `python famid_project_map.py <dir>` | famid → 計畫對照                           | `famid_project_map.csv` 等        |
| `python reconcile_famid.py <dir>`   | tool 資料夾 vs participants JSON 比對      | `orphan_*.csv`、`inactive_*.csv`  |
| `python scan_column.py <dir>`       | 找出含特定欄位關鍵字的檔案                 | `matched_files.xlsx`              |
| `python cell_extract.py <dir>`      | 讀資料夾內 xls 的指定儲存格整理成表        | `cell_extract.xlsx`               |
| `python cpt_convert.py [檔案]`      | 把 CPT 匯出的原始 txt 解析成 xlsx          | `cpt_parsed.xlsx`                 |
| `python tree.py <dir>`              | 匯出資料夾結構                             | `directory_tree.xlsx`            |

匯入、刪除、匯出、檢查類指令啟動後會互動選擇要處理的量表。

---

## 匯入 / 刪除 / 匯出

### 標準匯入（main.py）

1. 輸入計畫代碼（選填，按 Enter 跳過）
2. 互動選擇要上傳的量表
3. 讀取 Excel `import` sheet，過濾完全空白的列
4. 依 `config.py` 中的 `COLLECTION_MAP` 拆分欄位到各量表
5. 驗證每個欄位值（型別、範圍）
6. 量表欄位全為空值 / 999 的行自動跳過
7. 必要欄位（`famid` / `record_date` / `role`）檢查，缺少則報錯（見下方說明）
8. 驗證失敗的行寫入 `errors.xlsx`
9. 通過驗證的資料 upsert 至 MongoDB；若有輸入計畫代碼，每筆加上 `research_project_code` 欄位

#### 必要欄位檢查（famid / record_date / role）

分兩種狀況處理，缺少時顯示對應錯誤訊息：

1. **整欄不存在**（Excel 沒有該欄）→ 讀檔時即中止匯入，印出 `匯入中止：Excel 缺少必要欄位欄 {欄位名}`
2. **某列缺值**（有量表資料但缺其中一欄）→ 該列寫入 `errors.xlsx`，錯誤訊息為 `缺少必要欄位: {欄位名}`

> `record_date` 無法解析的值視同缺值，由上述「某列缺值」邏輯報錯；完全空白的列則直接略過，不視為錯誤。

### 年齡匯入（age_main.py）

適用於只有 `age`（歲數）但沒有確切 `record_date` 的資料：

1. 讀取 Excel，篩選有 `famid` + `age` 的行
2. 驗證量表欄位值（同標準匯入）
3. 從 `Participants` collection 查詢 `birth_date`，計算已存在紀錄的年齡
4. 若已存在紀錄的年齡與匯入的 age 相差 ≤ 1 歲 → 視為潛在重複，寫入 `no_date_records.xlsx`
5. 無衝突的資料：以 `birth_date + age` 估算 `record_date`，標記 `est_record_date: 1`，upsert 至 MongoDB
6. 每筆匯入資料自動加上 `raw_{scale}_age` 欄位保留原始年齡

### 刪除資料（delete_main.py）

依 `famid` + `record_date` 從選定的量表 collection 刪除資料。

### 匯出 JSON（export_main.py）

將 MongoDB 中的量表資料匯出為 JSON 檔案：

1. 互動選擇要匯出的量表（或全部）
2. 自動建立 `exports/YYYYMMDD/` 資料夾（以當日日期命名）
3. 每個量表匯出為獨立的 `{collection_name}.json`

---

## 統計 / 檢查

### Collection 統計（collection_stat.py）

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

### 單一 Collection 查詢（check_stat.py）

互動式輸入 collection 名稱，顯示總 document 數量，並依 `role` 分組計數。輸入 `q` 離開。

### 上傳前重複檢查（precheck_upload.py）

在**不寫入 MongoDB** 的前提下，拿一份新的 Excel 資料跟資料庫現有資料比對，用來 double-check「會不會重複上傳」以及「日期、計畫是否相符」。完全沿用 `main.py` 的流程（輸入計畫代碼 → 選擇量表 → `read_xlsx` → `split_by_collection`），所以檢查的就是「真的會被上傳」的那批資料，且不更動任何現有模組。

對資料庫**只有查詢（find）、沒有任何寫入**，唯一的輸出是本機報表 `precheck_report.xlsx`。

```bash
python precheck_upload.py              # 預設讀 import_data.xlsx
python precheck_upload.py 其他資料.xlsx  # 指定要檢查的檔案
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

### 欄位關鍵字掃描（scan_column.py）

遞迴掃描資料夾下的 `.xlsx` / `.xls` / `.xlsm` / `.csv` / `.tsv`，找出欄名含特定關鍵字（預設 `time` / `test` / `study`）的檔案與 sheet，輸出 `matched_files.xlsx`。

```bash
python scan_column.py <dir>
python scan_column.py <dir> -o out.xlsx --keywords date time --ignore .git venv
```

### 指定儲存格擷取（cell_extract.py）

掃描資料夾下所有 `.xls`（也支援 `.xlsx` / `.xlsm`），依「儲存格位址 → 輸出欄名」對應表，讀出每個檔案的特定儲存格，整理成一份新的 xlsx（每個檔案一列，第一欄附 `source_file` 來源檔名）。對應表由使用者自行定義：用 `--map A1=id B5=name`（行內，可多組）或 `--map-file`（JSON）指定。

```bash
python cell_extract.py <dir> --map A1=id B5=name C2=score   # 行內定義對應
python cell_extract.py <dir> --map-file mapping.json        # JSON：{"A1": "id", "B5": "name"}
python cell_extract.py <dir> --map A1=id --sheet Data       # 指定工作表（名稱或索引，預設第一個）
python cell_extract.py <dir> --map A1=id -r                 # 遞迴子資料夾（預設只掃最外層）
python cell_extract.py <dir> --map A1=id -o out.xlsx
```

- 超出範圍的儲存格留空、不報錯；整數型數值去除 `.0`、字串去頭尾空白
- 讀檔失敗的檔案仍保留該列，於 `_note` 欄記錄原因並標色

### CPT 原始檔解析（cpt_convert.py）

把 CPT 匯出的原始 txt 解析成 xlsx。每筆紀錄固定為 8 行的區塊，逐行拆出受測者資訊（自動偵測有無 `name` 欄）、用藥資訊與各項分數欄位，輸出單一 sheet（`CPT`）的 xlsx。

```bash
python cpt_convert.py                 # 預設讀 cpt_raw.txt，輸出 cpt_parsed.xlsx
python cpt_convert.py cpt_raw.txt     # 指定要解析的 txt
python cpt_convert.py cpt_raw.txt -o out.xlsx   # 指定輸出檔名
```

- `famid` / `famid_dup` 字串化並去除尾端 `.0`；數值欄自動轉型，全空的 `empty_` 欄自動丟棄
- 欄位順序整理為受測者資訊在前、分數欄位在後

### 資料夾結構匯出（tree.py）

將資料夾結構（含層級、類型、副檔名、大小、相對路徑）匯出為 Excel。

```bash
python tree.py <dir>
python tree.py <dir> -o directory_tree.xlsx --ignore .git node_modules
```

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

| 檔案                          | 說明                                                                                                               |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `errors.xlsx`                 | 驗證失敗的資料（每次執行新增 sheet）                                                                               |
| `no_date_records.xlsx`        | 年齡衝突或無法估算日期的資料（含 timestamp）                                                                       |
| `import_log.xlsx`             | **log** sheet：每次匯入的完整紀錄（含 success / insert / skipped_dup 欄位）；**pivot** sheet：依量表分組的統計數量 |
| `exports/YYYYMMDD/*.json`     | 從 MongoDB 匯出的量表 JSON 檔案                                                                                    |
| `exports/collection_stat/*.xlsx` | collection famid 統計（summary / by_last_digit / duplicates）                                                  |
| `precheck_report.xlsx`        | 上傳前重複檢查報表（唯讀產出）：Summary / Duplicates / Project_Mismatch / New_Rows / Date_Check                    |
| `compare_famid.xlsx`          | 兩份 participants JSON 比對結果                                                                                    |
| `famid_project_map.csv` / `project_summary.csv` | famid → 計畫對照與計畫摘要                                                              |
| `orphan_famid_report.csv` / `inactive_participant_report.csv` | tool vs participants 雙向比對結果                                        |
| `matched_files.xlsx`          | 含特定欄位關鍵字的檔案清單                                                                                         |
| `cell_extract.xlsx`           | 資料夾內各 xls 指定儲存格的擷取結果（每檔一列）                                                                    |
| `cpt_parsed.xlsx`             | CPT 原始 txt 解析結果（單一 `CPT` sheet）                                                                          |
| `directory_tree.xlsx`         | 資料夾結構                                                                                                          |

---

## 支援量表

GSQ、SNAP4、CBCL、AQ、SRS、SAICA、SCQ、CAST、BRIEF、CPRS、CTRS、SDQ、PMI、MPBI、DPBI、APGAR、EPQ、YSR、SUB、ASRI、ASRS、CES-D、MPI、RBS-R、SSP、WFIRS-P、WFIRS-S、WHOQOL、CEQ、C-SBEQ、ERQ-CA、BRIEF-S、BRIEF-A、ERQ-A、RAADS-R、TAS-20、AAQOL、ARI、ESQ
