# backup_iq 四檔 + datapool 大表分析結論與整理方案

> 分析日期：2026-07-11
> 資料來源：`C:\Users\user\Downloads\rawdata\iq\backup_iq`（4 個 xlsx，取 `raw` worksheet）＋ `C:\Users\user\Downloads\rawdata\iq\datapool`（3 個統整大表）
> 本文件僅為分析結論與整理方案，未寫入任何程式或 DB（DB 僅唯讀查詢比對）。
> ⚠ 重要前提：**DB 中已存在 `IQ` collection（2,075 筆）與 `temp_iq`（594 筆）**，由本 repo 以外的管道匯入（repo 內 `iq_import.py` 為空檔、無匯入紀錄）。詳見第 6–7 節。

---

## 1. 背景與檔案清單

原檔皆未標示 IQ 測驗種類（WISC / WPPSI / WAIS…），需先判定版本才能對應欄位。現況：

- `iq/iq_rename.py`、`iq/iq_import.py` 為 0-byte 空檔，IQ 匯入管線尚未建置；僅 `iq/iq_precheck.py`（key = famid + age，資料驅動比對）可用。
- **IQ codebook 其實已存在**：`TD_IQ_20221102_nsc98_adhd.xlsx` 的 `code` worksheet 完整定義了「統一後變數」（snake_case）、變數型態（`TEXT` / `INT` / `FLOAT` / `DATE`）、各版本原始欄名對照（WISC_III / WISC_IV / WISC_V / WAIS_III(+nhri100_adult 變體) / WAIS_IV / WPPSI_R / WPPSI_IV），以及 `version` 欄位的合法詞彙：`WISC_V, WISC_IV, WISC_III, WAIS_III, WAIS_IV, WPPSI_R, WPPSI_IV, SPM, CPM, SB5, other(其他來源 e.g. 之前做過/家長提報告)`。

| 檔案 | raw 列數 | 欄數 | 推測計畫代碼（檔名） |
| --- | --- | --- | --- |
| `ADHD_TW+NHRI+FU_IQ_nhri94.xlsx` | 618 | 37 | nhri94 |
| `MOST107IQ_200907_most107.xlsx` | 358 | 15 | most107 |
| `TD_IQ_20221102_nsc98_adhd.xlsx` | 116 | 35 | nsc98_adhd |
| `WPPSI221012_nsc96_asd.xlsx` | 227 | 31 | nsc96_asd |

---

## 2. 版本判定與證據

判定方法：以 code sheet 的各版本欄名對照為基準，用「分測驗欄位簽名」反推版本。結論：

| 檔案 | 判定 | 決定性證據 |
| --- | --- | --- |
| nhri94 | **WISC_III**（待人工確認） | 有 `FDI`（專心注意指數）——codebook 中**僅 WISC_III** 有此指數；並有 `VIQ`/`PIQ`（WISC_IV 起取消）與 `PA` 連環圖系、`OA` 物型配置（WISC_IV 無）。測驗年民國 91–101、年齡 2–20 歲，符合台版 WISC-III 使用年代 |
| nsc98_adhd | **WISC_III**（待人工確認） | 同上 signature（FDI + VIQ/PIQ + PA/OA）。測驗年 2017–2020 仍用 WISC_III 需 postdoc 確認（縱貫一致性？） |
| nsc96_asd | **WPPSI_R** | `Sen`（句子）、`GF`（幾何圖形）——codebook 中僅 WPPSI_R 有；檔名亦標 WPPSI |
| most107 | **無法判定** | 僅有總分（VIQ/PIQ/FIQ/VCI/POI/FDI/Z_score），無任何分測驗欄位。VCI/POI/FDI 組合為 WISC_III 時代計分結構，但無法排除混合來源；可能屬 `other`（先前施測 / 家長提供報告） |

注意：nhri94 與 nsc98 雖同為 WISC_III 內容，**欄名縮寫是兩套不同系統**（nhri94 用 `IN/SI/AR/VC/CO/PCm/Cc`＝WISC_IV 風格縮寫，顯示曾被人工改名過；nsc98 用 `In/Ss/A/V/C/Ds/DS_f/DS_b`）。rename 對照必須「一檔一設定」，不能只做一張全域對照表。

依決議：`version` 欄位**匯入時留 null**，本文件證據供 postdoc 逐檔確認後再回填。

---

## 3. 逐檔資料品質問題

以下數字皆經腳本實跑驗證。

### 3.1 nhri94（618 列）

- **日期為民國年**：出生 `cyear` 69–81、測驗 `year` 91–101。repo 目前**沒有任何民國年轉換程式**，需新寫（+1911）。
- **999 sentinel 普遍**：測驗日期 999 有 20 列、出生日期 999 有 2 列、`age_IQ`=999 有 20 列；各分測驗 999 約 22–23 列、`LDSF`/`LDSB` 各 35 列。⚠ 有 20 列「無測驗日期且 age_IQ 也是 999」→ 匯入時屬 `flag_no_date` 類。
- **famid 長度 2–5 碼**（fam 1–4 碼 + id 1 碼），`Fam` 首碼以 `1` 為大宗（533/618）→ 即 CLAUDE.md 待決事項 #3 的「`1`–`169`、`1001`–`1324` 屬 TD 或 ADHD 不明」群，本檔無法解套此問題。
- `case` 欄有 0/1/2/5 四種值，意義未見文件定義；依核心原則屬時變欄位，保留原值待確認。
- `id` 僅 1（個案）與 4（手足）。

### 3.2 most107（358 列）——建議整批暫緩

- **完全沒有測驗日期，也沒有測驗年齡**（僅出生年月日與總分）→ composite key 無法用 `famid+record_date` 也無法用 `famid+age`。
- 999 sentinel：`VIQ` 111 列、`Z_score` 112 列、`FIQ` 8 列；**7 列六項總分全為 999、另 68 列 IQ 欄位全空**（有 famid 但無任何 IQ 資料）。
- `Z_score` 欄意義不明（值域 42–135 左右，較像標準分數而非 z 分數），待 postdoc 釐清。
- famid `417321` 重複出現 2 列。
- `site` 含未定義代碼 `7`（10 列）——CLAUDE.md 僅定義 1=台大、2=長庚、3=市療、4=桃療。
- **決議：整批暫緩不進 DB**，先由 postdoc 確認資料來源（疑為先前施測/家長提供報告）後再決定匯入方式。

### 3.3 nsc98_adhd（116 列）

- **測驗年錯字**：famid 87461 的 `year=207`、famid 87721 的 `year=1018`（推測為 2017 / 2018，但依核心原則不自動修正，匯出人工確認）。
- **超範圍值**：famid 87071 的 `DS_f=117`、`DS_b=87`（量表分數合法上限約 19，疑為串接輸入）。
- 檔內另有 `work`（幾乎同 raw）與 `code`（= IQ codebook，見第 1 節）兩個 worksheet；資料以 `raw` 為準。
- famid 全為 87xxx → TD 組（87–89 開頭 = TD），與檔名 `TD_IQ` 一致。

### 3.4 nsc96_asd（227 列）

- `fam1` / `id1` 兩欄 226/227 列為 999（近乎垃圾欄），但**有 1 列 fam1=73009 / id1=4**——7 開頭疑為 nsc99 編號系統，需浮現不可靜默丟棄。
- 1 列測驗 `year=999`（無測驗日期）。
- 999 sentinel：`Sen` 65 列、其餘分測驗各 13–16 列、`PIQ` 9 / `VIQ` 7 / `FIQ` 1 列。
- **`site` 含未定義代碼 `5`（41 列）**、另 1 列 site=3；CLAUDE.md site 對照僅定義 1–4。
- `id` 含 5（3 列）＝第二位手足，屬合法值（4+ 依序遞增）。
- famid 為六碼 ASD 格式（`4`+site+fam+id），全檔一致。

### 3.5 跨檔關係

- **most107 ∩ nsc96_asd famid 重疊 35 人**（如 410301、411441…）→ 縱貫重測資料，依核心原則保留所有記錄、絕不自動合併。
- 其餘檔案兩兩之間 famid 無重疊。

---

## 4. 建議整理方案（未來管線，本輪不實作）

> 註：DB 已有 2,075 筆 IQ 資料（見第 6 節），因此匯入步驟的重點從「首次匯入」變成「**與現有 DB 比對後補匯**」——exact_dup 跳過、value_conflict 出報告不覆蓋，正好是 `src/importer.py` upsert 分類邏輯的用途。

整體仿 CANTAB 管線四件套（`cantab/cantab_config.py` → `cantab_rename.py` → `cantab_precheck.py` → `cantab_import.py`），並重用 CLAUDE.md 第 8 節的 src/ 共用模組（`get_db()`、`wait_and_retry`、各 writer）。

1. **`iq/fields/`**：把 nsc98 檔的 `code` sheet 轉成正式 codebook——
   - `{日期}_IQ_fields.json`：統一後變數 → `<INT>[lo,hi]` / `<FLOAT>` / `<DATE>` 型別規格（沿用 `parse_field_spec` 格式）；
   - rename 對照表 xlsx：一版本一欄（WISC_III / WPPSI_R / …），且同一版本要能容納多套縮寫（nhri94 與 nsc98 同為 WISC_III 但欄名不同）。
2. **`iq/iq_config.py`**：`IQ_COLLECTION = "IQ"`、`IQ_SHARED_FIELDS = ["famid", "record_date"]`、version 合法詞彙表，以及**每來源檔一筆 SOURCE_CONFIG**：version 建議值（供報告，不直接寫入）、`research_project_code`（nhri94 / most107 / nsc98_adhd / nsc96_asd）、曆制（民國/西元）、垃圾欄清單（如 nsc96 的 fam1/id1）、日期欄組合（cyear+cmonth+cday → birth_date；year+month+day → record_date）。
3. **`iq/iq_rename.py`**：欄名 → 統一後變數（附錄 A 對照）、民國年 +1911 轉西元、famid 清理（str + 去 `.0`）、**999 → null（IQ 分數 999 為 sentinel，此轉換僅限 IQ 管線**，不可外溢到 ADI-R）、未知欄位保留原名標黃浮現（同 cantab_rename 規則 4）。
4. **`iq/iq_precheck.py`**（既有）：跑清理後檔案，產出重複/衝突/缺 key 報告。
5. **`iq/iq_import.py`**（仿 cantab_import）：
   - 有 record_date → 正常以 `famid+record_date` 匯入；
   - 無日期但有年齡 → `record_date=null` + `age_suspect=true`＋保留原始年齡欄；
   - 無日期無年齡 → `record_date=null` + `flag_no_date=true`；
   - `version` 本輪一律 null，另出「版本確認清單」xlsx（附本文件第 2 節證據）給 postdoc；
   - 全部文件加 `research_project_code`。
6. **錯誤浮現**：nsc98 的年份錯字與超範圍 DS_f/DS_b、nsc96 的 fam1=73009 等 → 進 error report（`src/error_writer.py`），不自動修。

## 5. datapool 三大表分析

`datapool/` 為另外統整過的大表，一檔一組別、一 sheet 一測驗家族，**已有 `Version` 欄位**（文字），欄名大致等同 codebook 英文代碼。實際資料列數（去除表頭雜列後）：

| 檔案 | sheet | 列數 | Version 分布 |
| --- | --- | --- | --- |
| IQ_data_ADHD.xlsx | WISC | 87 | WISC_V 66、WISC_IV 17、**DK 4** |
| IQ_data_ADHD.xlsx | WPPSI | 27 | WPPSI_IV 25、**DK 2** |
| IQ_data_ASD_250616.xlsx | WPPSI | 152 | WPPSI_IV 143、WPPSI_R 7、SB5 1、**999 1** |
| IQ_data_ASD_250616.xlsx | WISC | 204 | WISC_V 90、WISC_IV 85、WISC_III 1、CPM-P 1、**999 27** |
| IQ_data_ASD_250616.xlsx | WAIS | 19 | WAIS-IV 16、WAIS-III 3 |
| IQ_data_TD_20251113.xlsx | WPPSI | 69 | WPPSI_IV 69 |
| IQ_data_TD_20251113.xlsx | WISC | 207 | WISC_V 162、**DK 45** |
| IQ_data_TD_20251113.xlsx | WAIS | 64 | WAIS_IV 57、WAIS_3 7 |

結構問題：

- **雙列表頭**：ASD 的 WPPSI sheet 表頭在第 2 列（第 1 列為 Unnamed/中文）；ADHD 與 TD 的 WPPSI sheet 分測驗欄第 1 資料列是中文標籤列。讀取時需先剝除。
- **Version 字串不一致**：`WAIS-III` / `WAIS_3` / `WAIS_IV` / `WAIS-IV` 混用，另有 `DK`（51 列）與 `999`（28 列）待確認；`CPM-P`、`SB5` 不在 codebook 詞彙中（CPM-P 近似 CPM）。
- **ADHD、TD 兩檔完全沒有 `record_date`**（只有 age_y/m/d）；ASD 檔有 record_date 但含 999。
- 檔內重複 famid：ADHD/WISC 37161×2、ADHD/WPPSI 37011×2（疑重測，保留）。
- famid 組別與前綴一致：ADHD=3 開頭、TD=5 開頭、ASD=4 開頭六碼，皆符合編碼規則。

### 與 backup_iq 的重疊

- pool_ASD ∩ most107：**33 famid**；pool_ASD ∩ nsc96：3 famid（413694、415854、440294）。
- 其餘組合皆為 0。→ backup 四檔（舊年代）與 datapool（新年代）幾乎互補，datapool **不能取代** backup 檔。

## 6. 與 DB `IQ` collection 比對結果（唯讀查詢）

DB.IQ 現況：2,075 docs；欄位**已使用 codebook 統一後變數**（information_score…）；`record_date` 為 YYYY-MM-DD 文字（僅 13 筆空）；**但 `version` 存的是數字代碼 1–8**，repo 內找不到對照表。依欄位特徵與年代反推：

| version 代碼 | docs | record_date 年代 | 特徵欄位 | 推測版本 |
| --- | --- | --- | --- | --- |
| 1 | 594 | 2002–2007 | FDI、VIQ/PIQ、PA | **WISC_III**（= nhri94 那批，數量吻合） |
| 2 | 99 | 2014–2023 | MR、Cd、WMI | WISC_IV？ |
| 3 | 308 | 2020–2026 | VSI、FRI、LDSS | WISC_V |
| 4 | 734 | 2006–2020 | VIQ、PA、MR、WMI | WAIS_III？ |
| 5 | 73 | 2019–2023 | MR、Sq、WMI、LDSS | WAIS_IV |
| 6 | 214 | 2014–2025 | BS 昆蟲尋找、ZL 動物園 | WPPSI_IV |
| 7 | 46 | 2019–2022 | MR、WMI、VSI、FRI | ?（WISC_V 類） |
| 8 | 7 | 2013–2021 | 六碼 famid、VIQ、MR | ?（SB5/other？） |

`temp_iq`（594 筆、欄位同 IQ 但較少）疑為 nhri94 匯入時的暫存 collection，是否清除待決。

### 6.1 backup_iq 四檔 vs DB（記錄層級，key = famid + record_date）

| 檔案 | key 完全對上 | 其中欄位值衝突 | 未入庫 |
| --- | --- | --- | --- |
| nhri94（民國轉西元後比對） | **593/618** | 全欄位比對後**真差異僅 2 筆**；另發現系統性欄位對調與 sex 缺漏（見下） | 25 列（多為無測驗日期 999 者） |
| most107 | 0 | — | 全數未入庫（326/358 連 famid 都不在 DB） |
| nsc98_adhd | 0 | — | **全數未入庫**（116 列 famid 皆不在 DB） |
| nsc96_asd | 0 | — | 幾乎全未入庫（僅 3 famid 出現在 DB，屬 datapool 來源） |

⚠ **nhri94 全欄位比對的三個發現**（詳見 `output/20260711_IQ_nhri94_conflict_report.xlsx`）：

1. **backup 檔六個總分欄位標頭錯置**：檔案與 DB 之間 VIQ↔VCI、PIQ↔POI、FIQ↔FDI（含各自百分等級）成對互換，共 1,652 筆欄位值。內部一致性檢查（全量表智商應介於 VIQ 與 PIQ 之間）顯示 DB 標頭下平均偏差 1.04、檔案標頭下 3.64 → **DB 的值是對的，backup 檔的欄位標頭標錯**。初次比對誤判的「144 筆 FIQ 衝突」即源於此。
2. **DB 該批缺 sex**：檔案有 sex 但 DB 為空共 592 筆（另 raw_iq_age 亦有大量缺漏/差一歲）→ 可由檔案回補（合計「檔案有值、DB 為空」1,623 筆欄位值）。
3. 對調校正後**真差異僅 2 筆**、「DB 有值檔案 999」11 筆（famid 264 整列分測驗）→ DB 匯入來源與 backup 檔實為同一批資料。

### 6.2 datapool vs DB（ADHD/TD 無日期，改以 famid + FIQ + age_y 比對）

| sheet | 已在 DB（完全一致） | 未入庫 |
| --- | --- | --- |
| pool_ADHD/WISC | 86/87 | 1 |
| pool_ADHD/WPPSI | 23/27 | 4 |
| pool_TD/WPPSI | 62/69 | 7 |
| pool_TD/WISC | 201/207 | 6 |
| pool_TD/WAIS | 63/64 | 1（famid 在 DB 但 FIQ 不同，衝突或重測） |
| pool_ASD/WPPSI | 136/152 | 16 |
| pool_ASD/WISC | 166/204 | 38 |
| pool_ASD/WAIS | 19/19 | 0 |
| **合計** | **756/829（91%）** | **73** |

- datapool **九成一已在 DB**（FIQ 與 age_y 完全一致，且 pool_ASD 的 182 筆連 famid+record_date 都完全對上）→ DB 的新世代資料極可能就是由 datapool（或其上游）匯入。
- 值得注意：pool_ADHD / pool_TD 檔內沒有 record_date，但 DB 對應 docs **有**日期 → DB 匯入來源比 datapool 大表更完整；datapool 是衍生表，不是最上游。
- 未入庫的 73 列是後續補匯目標（多為 ASD 新個案，54 列）。

## 7. 待 postdoc 決策清單

| # | 事項 | 相關檔案 |
| --- | --- | --- |
| 1 | most107 整批資料來源確認（是否為先前施測/家長提報告）與是否匯入、如何定 key | most107 |
| 2 | version 確認：nhri94 / nsc98 是否確為 WISC_III；most107 的 version 值 | nhri94, nsc98, most107 |
| 3 | site 代碼 `5`（41 列）與 `7`（10 列）的意義（現行文件僅定義 1–4） | nsc96, most107 |
| 4 | nhri94 famid（1 開頭家庭）屬 TD 或 ADHD（= CLAUDE.md 待決事項 #3） | nhri94 |
| 5 | nsc98 測驗年錯字修正值（207→2017？1018→2018？）與 DS_f=117/DS_b=87 處理 | nsc98 |
| 6 | nhri94 `case` 欄（0/1/2/5）意義 | nhri94 |
| 7 | most107 famid 417321 重複 2 列、68 列全空記錄是否保留 | most107 |
| 8 | nsc96 fam1=73009（7 開頭疑 nsc99 編號）該列的身分對應 | nsc96 |
| 9 | **DB.IQ `version` 數字代碼 1–8 的正式對照表**（repo 無定義；第 6 節有反推）；後續匯入要沿用數字碼還是改 codebook 文字詞彙 | DB |
| 10 | **nhri94 backup 檔六個總分欄位標頭錯置**（VIQ↔VCI、PIQ↔POI、FIQ↔FDI；一致性檢查顯示 DB 值正確）：確認後修正 backup 檔標頭或註記作廢；DB 缺漏的 sex（592 筆）是否由檔案回補；真差異 2 筆判定 | nhri94, DB |
| 11 | `temp_iq` collection（594 筆，疑 nhri94 匯入暫存）是否清除 | DB |
| 12 | datapool 未入庫 73 列的補匯；其中 Version=DK 51 列、999 28 列如何標記；`WAIS-III`/`WAIS_3` 等字串正規化 | datapool |
| 13 | DB.IQ 由 repo 外管道匯入——找回當初的匯入腳本/來源檔，納入版控或在本 repo 重建 | DB |

## 附錄 A：各檔 raw 欄位 → 統一後變數對照草表

依 code sheet 統一後變數；「（不進 DB）」= famid 組成元件或垃圾欄，待確認欄位以 ⚠ 標示。

### A.1 nhri94（WISC_III 待確認；日期為民國年）

| raw 欄 | 統一後變數 | 備註 |
| --- | --- | --- |
| famid | famid | Fam、id 為組成元件（不進 DB） |
| case | ⚠ 待定 | 時變 group 欄，意義待確認 |
| sex | sex | |
| age_IQ | age_y | 測驗年齡_年；999→null |
| cyear/cmonth/cday | birth_date | 民國+1911 → YYYY-MM-DD |
| year/month/day | record_date | 民國+1911 → YYYY-MM-DD；999→null+flag |
| IN | information_score | |
| SI | similarity_score | |
| AR | arithmetic_score | |
| VC | vocabulary_score | |
| CO | comprehension_score | |
| Dsp | memory_span_score | |
| PCm | picture_completion_score | |
| Cc | symbol_substitution_score | |
| PA | picture_arrangement_score | |
| BD | figure_design_score | |
| OA | object_assembly_score | |
| LDSF | longest_forward_span | |
| LDSB | longest_backward_span | |
| VIQ / VIQ_P | verbal_scale_iq / verbal_scale_iq_percentile | |
| PIQ / PIQ_P | performance_scale_iq / performance_scale_iq_percentile | |
| FIQ / FIQ_P | full_scale_iq / full_scale_iq_percentile | |
| VCI / VCI_P | verbal_comprehension_index / verbal_comprehension_percentile | |
| POI / POI_P | perceptual_organization_index / perceptual_organization_percentile | |
| FDI / FDI_P | focused_attention_index / focused_attention_percentile | |

### A.2 nsc98_adhd（WISC_III 待確認；西元）

分測驗縮寫不同、其餘同 A.1：

| raw 欄 | 統一後變數 |
| --- | --- |
| Famid | famid（Fam、Id 不進 DB） |
| In | information_score |
| Ss | similarity_score |
| A | arithmetic_score |
| V | vocabulary_score |
| C | comprehension_score |
| Ds | memory_span_score |
| PC | picture_completion_score |
| Cd | symbol_substitution_score |
| PA / BD / OA | 同 A.1 |
| DS_f | longest_forward_span |
| DS_b | longest_backward_span |
| VIQ…FDI_P | 同 A.1 |

### A.3 nsc96_asd（WPPSI_R；西元）

| raw 欄 | 統一後變數 | 備註 |
| --- | --- | --- |
| famid | famid | site/family/id/fam 為組成元件（不進 DB） |
| fam1 / id1 | （不進 DB） | 垃圾欄；⚠ 惟 fam1=73009 該列需浮現 |
| sex | sex | |
| cyear/cmonth/cday | birth_date | |
| year/month/day | record_date | 999→null+flag |
| OA | object_assembly_score | |
| GF | geometry_score | WPPSI_R 專屬 |
| BD | figure_design_score | |
| MR | matrix_reasoning_score | |
| PC | picture_completion_score | |
| In | information_score | |
| C | comprehension_score | |
| A | arithmetic_score | |
| V | vocabulary_score | |
| Ss | similarity_score | |
| Sen | sentence_score | WPPSI_R 專屬；999 有 65 列 |
| PIQ/VIQ/FIQ 及 _P | 同 A.1 | |

### A.4 most107（暫緩，僅供參考）

| raw 欄 | 統一後變數 | 備註 |
| --- | --- | --- |
| famid | famid | site/fam/id 不進 DB |
| csex | sex | |
| cyear/cmonth/cday | birth_date | |
| VIQ/PIQ/FIQ/VCI/POI/FDI | 同 A.1 各 index | 999→null |
| Z_score | ⚠ 待定 | 意義不明，待 postdoc |

（無 record_date、無年齡欄位）
