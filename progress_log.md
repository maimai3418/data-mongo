# SGLAB 資料整合進度記錄

> 最後更新：2026-07-07
> 2026-07-07 已對照實際程式碼逐一稽核，取代先前僅依 CLAUDE.md/commit message 推斷的版本。狀態欄位以「程式碼實際能做到什麼」為準，不採信文件或先前記錄裡未經查證的說法。

## 進度總覽

| 工具         | 階段                               | 狀態 | 最後更新   | 備註                                                                                                   |
| ------------ | ---------------------------------- | ---- | ---------- | ------------------------------------------------------------------------------------------------------ |
| ADOS         | 匯入管線完成                       | ✅   | 2026-07-07 | config / import / delete / field_rename 齊全；famid 兩種 fallback 已確認存在於程式碼                    |
| ADI-R        | 匯入管線完成                       | ✅   | 2026-07-07 | 999 例外邏輯（月齡欄位）與 CLAUDE.md §7/§9 描述一致，已核對程式碼                                        |
| **CPT**      | **匯入管線完成（三層去重複）**     | ✅   | 2026-07-07 | 新增 `cpt/cpt_config.py`（field.json 解析）+ `cpt/cpt_import.py`：famid+record_date 三層去重複（batch hard dedup / retest 保留 / 與 DB 衝突匯出人工審查）、no-date 列以 `flag_no_date` 匯入並記錄於 `no_date_records.xlsx`。已用合成資料 + mock DB 驗證邏輯正確；尚未對真實 DB / 真實檔案跑過。`cpt_precheck.py` 仍只比對批次內彼此，不查 DB（未改動） |
| **CANTAB**   | **匯入管線完成（三層去重複 + age_suspect）** | ✅   | 2026-07-07 | 新增 `cantab/cantab_config.py` + `cantab/cantab_import.py`：同 CPT 三層去重複；no-date 列依是否有年齡分流為 `age_suspect`+`raw_cantab_age`（首次落實 CLAUDE.md §9 的 age_suspect 慣例）或 `flag_no_date`。已用合成資料 + mock DB 驗證。**TXT→xlsx 轉換腳本仍不存在**（範圍外，未看過原始檔案格式）；尚未對真實 DB / 真實檔案跑過 |
| **IQ**       | **precheck（批次內去重）完成；匯入/重命名尚未撰寫** | 🔧   | 2026-07-07 | 原一次性 ad hoc script 已改名保留為 `iq/iq_test.py`（未刪除）。新 `iq_precheck.py` 比照 CPT/CANTAB 的通用批次比對工具：famid+age key，field.json 存在則採用，不存在則退回「資料驅動」模式（比對欄位＝本次讀到的欄位）。`iq_import.py`/`iq_rename.py` 仍是 0 行空檔 |
| Participants | 新增/precheck 腳本建立             | 🔧   | 2026-07-06 | `p_add_new.py`、`p_precheck.py` 存在（本輪未逐行稽核）                                                   |
| KE           | 尚無任何程式碼                     | 📋   | 2026-07-07 | repo 中無 `ke/` 資料夾；CLAUDE.md 所稱「欄位映射完成」在此 repo 找不到對應檔案                            |
| MRI          | 尚無任何程式碼                     | 📋   | 2026-07-07 | repo 中無 `mri/` 資料夾；CLAUDE.md 所稱「mongosh 聚合腳本」在此 repo 找不到對應檔案                       |
| Identity     | 未實作                             | ❌   | 2026-07-07 | `aliases`/`Identity` 全 repo 零命中，無任何腳本讀寫此 collection，純文件構想                             |
| famid_timepoint_map | 未動工                      | 📋   | 2026-07-07 | CLAUDE.md §6 有完整演算法描述，repo 中無此檔案；已於 2026-07-07 由 Agent 1 產出實作規格（見下方）        |

## 狀態圖例

- ✅ 管線完成（可能仍有後續維護）
- 🔧 進行中（有明確的下一步）
- 📋 早期/規劃階段
- ⏸️ 等待外部決策
- ❌ 阻塞 / 未實作

## 2026-07-07 稽核發現（與 CLAUDE.md / 舊 progress_log 的落差）

| 項目 | CLAUDE.md／舊記錄說法 | 實際程式碼狀態 |
| --- | --- | --- |
| CPT 匯入 | §9「🔧 去重複邏輯完成」 | `cpt_import.py` 空檔；precheck 不查 DB；三層去重複只有兩層 |
| CANTAB | §9「🔧 TXT→xlsx 轉換完成」 | 無轉換腳本；`cantab_import.py` 空檔；`age_suspect` 未實作 |
| IQ | §9「🔧 Codebook 完成」，舊記錄「precheck+rename+匯入完成」 | 三支腳本中兩支（import/rename）空檔；precheck 是寫死路徑的一次性 script |
| famid_timepoint_map.py | §6 完整演算法描述 | 完全不存在 |
| KE | §9「🔧 欄位映射完成」 | repo 無對應資料夾/程式碼 |
| MRI | §9「有 mongosh 聚合腳本」 | repo 無對應資料夾/程式碼 |
| Identity/aliases | §2、§9：45–77 對雙重 famid | 全 repo 零命中，無程式碼讀寫 |
| 五種類匯入分類（§4） | 隱含全面實作 | 通用管線只做 2 桶（new_insert / 跳過重複），不區分 code_mismatch vs value_conflict；只有 `precheck_upload.py` 會查 DB，但也不分兩種衝突桶 |

確認一致（程式碼確實落實文件規則）：famid 前綴分組（`collection_stat.py`）、ADI-R 999 例外、ADOS famid 兩種 fallback、日期存為文字、.env/get_db() 連線慣例。

程式碼存在但 CLAUDE.md 未記錄：`reconcile_famid.py`、`famid_compare.py`、`famid_project_map.py`、`collection_stat.py`、`check_stat.py`、`scan_column.py`、`cell_extract.py`、`sheet_compare.py`、`precheck_upload.py`（唯一會實際查 DB 做比對的通用工具）。

## 待決事項 (Pending Decisions)

| #   | 事項                                              | 需要誰決策 | 提出日期 | 影響範圍     |
| --- | ------------------------------------------------- | ---------- | -------- | ------------ |
| 1   | KE 舊版 collection 命名（`KE_V1`/`KE_V2` 或其他） | Postdoc    | 待確認   | KE 匯入管線  |
| 2   | 48 位新參與者缺少 `group` 欄位                    | Postdoc    | 待確認   | Participants |
| 3   | famid 前綴 `1`–`169`、`1001`–`1324` 究竟屬於 TD 或 ADHD | Postdoc | 待確認   | 全域         |
| 4   | MRI 資料範圍釐清                                  | Postdoc    | 待確認   | MRI_INFO     |

## 下一步優先順序（使用者指定：CPT → CANTAB → IQ，其餘不急）

### 1. CPT（已完成，2026-07-07）
- ✅ `cpt_import.py` 撰寫完成：famid+record_date 三層去重複、與 DB 衝突匯出人工審查、no-date 列標記匯入。
- 待辦（非阻塞）：`cpt_precheck.py` 仍只比對「同批次新檔案彼此」，不查 DB；若要滿足 CLAUDE.md §4「與 DB 現有記錄比對」，可另外補上，或明確定位 precheck 只負責批次內去重、DB 比對交給 import 腳本（目前採後者的分工）。
- 待辦（非阻塞）：尚未用真實 CPT 檔案 + 真實 MongoDB 連線做端對端驗證（本次僅用合成資料 + mock DB 驗證邏輯）。

### 2. CANTAB（已完成，2026-07-07；一項待辦仍未解決）
- ✅ `cantab_import.py` 撰寫完成：同 CPT 三層去重複，另補上 `age_suspect`/`raw_cantab_age` 邏輯。
- 待辦（阻塞 TXT 原始檔匯入）：TXT→xlsx 轉換腳本仍不存在。若 CANTAB 原始資料確實是 txt（CLAUDE.md §9 暗示），需要先看過實際原始檔案格式才能撰寫轉換腳本；目前 `cantab_import.py` 只能吃 xlsx。
- 待辦（非阻塞）：尚未用真實 CANTAB 檔案 + 真實 MongoDB 連線做端對端驗證。

### 3. IQ（precheck 已完成，import/rename 待撰寫）
- ✅ 舊一次性腳本改名保留為 `iq_test.py`；新 `iq_precheck.py` 為通用批次比對工具（同 CPT/CANTAB 風格，field.json 選用）。
- 待辦：沒有 `iq/fields/*.json`（無 codebook），目前 precheck 靠資料驅動模式運作。若日後拿到正式 codebook，放入 `iq/fields/` 即會自動切換使用。
- 待辦：撰寫 `iq_import.py`、`iq_rename.py`（目前皆 0 行），比照 CPT/CANTAB 補上三層去重複 + DB 比對。

### 其他（不急，維持現狀記錄）
- `famid_timepoint_map.py`：Agent 1 已於 2026-07-07 產出實作規格（相鄰日期聚類、三分類、兩階段 report/apply、_id 為更新 key），尚未動工。
- KE / MRI：repo 中無程式碼，待資料/決策到位。
- Identity/aliases：概念存在但無實作，暫緩。

## 已產出的關鍵檔案

| 檔案                        | 類型     | 用途                                          |
| --------------------------- | -------- | --------------------------------------------- |
| ados/ados_config.py         | 腳本     | ADOS 欄位設定                                 |
| ados/ados_field_rename.py   | 腳本     | ADOS 欄位重命名（含 famid 兩種 fallback）      |
| ados/ados_import.py         | 腳本     | ADOS 匯入                                      |
| ados/ados_delete.py         | 腳本     | ADOS 刪除                                      |
| adir/adir_config.py         | 腳本     | ADI-R 欄位設定（999 例外邏輯）                 |
| adir/adir_field_rename.py   | 腳本     | ADI-R 欄位重命名                               |
| adir/adir_import.py         | 腳本     | ADI-R 匯入                                     |
| cpt/cpt_convert.py          | 腳本     | CPT 原始 txt → xlsx                            |
| cpt/cpt_precheck.py         | 腳本     | CPT 批次內去重（不查 DB）                      |
| cpt/cpt_config.py           | 腳本     | CPT field.json 解析（2026-07-07 新增）         |
| cpt/cpt_import.py           | 腳本     | CPT 匯入：三層去重複 + DB 比對（2026-07-07 完成） |
| cantab/cantab_precheck.py   | 腳本     | CANTAB 批次內去重（不查 DB）                   |
| cantab/cantab_config.py     | 腳本     | CANTAB field.json 解析（2026-07-07 新增）      |
| cantab/cantab_import.py     | 腳本     | CANTAB 匯入：三層去重複 + age_suspect（2026-07-07 完成） |
| iq/iq_precheck.py           | 腳本     | IQ 批次內去重（不查 DB，同 CPT/CANTAB 風格，field.json 選用，2026-07-07 重寫） |
| iq/iq_test.py               | 一次性腳本 | 原 iq_precheck.py 改名保留：寫死路徑的單次 famid+IQ 分數比對 |
| iq/iq_import.py             | 空檔     | 尚未撰寫                                       |
| iq/iq_rename.py             | 空檔     | 尚未撰寫                                       |
| participants/p_add_new.py   | 腳本     | 新增參與者                                     |
| participants/p_precheck.py  | 腳本     | 參與者資料上傳前檢查                            |
| precheck_upload.py          | 腳本     | 唯一會查 DB 比對現有紀錄的通用上傳前檢查工具    |
| reconcile_famid.py          | 腳本     | tool vs participants famid 雙向比對             |
| famid_compare.py            | 腳本     | 兩份 participants JSON famid 比對               |
| famid_project_map.py        | 腳本     | famid → 專案映射                                |
| collection_stat.py          | 腳本     | famid 前綴分組統計（與 CLAUDE.md §2 規則一致）  |

## 建議的未來驗證方向

- 抽樣審計：抽取紙本原件比對，偵測範圍內的轉錄錯誤
- 跨工具時序一致性：同 famid 在不同工具的施測日期是否合理
- CLAUDE.md 本身有多處敘述領先於實際程式碼（見上方稽核發現），建議定期（每次重大進度後）重新核對，避免文件與現況繼續脫節
