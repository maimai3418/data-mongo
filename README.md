# data-mongo

將心理衡鑑量表資料（xlsx）依量表類型拆分，驗證後匯入 MongoDB。

---

## 資料夾結構

```
data-importer/
├── .env
├── .gitignore
├── requirements.txt
├── config.py
├── main.py
├── delete_main.py
└── src/
    ├── __init__.py
    ├── reader.py
    ├── transformer.py
    ├── importer.py
    ├── error_writer.py
    └── logger.py
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

將 `.xlsx` 檔案放入專案根目錄，並在 `main.py` 修改檔名：

```python
FILEPATH = "import_data.xlsx"
```

### 4. 執行匯入

```bash
python main.py
```

---

## 說明

- 支援量表：GSQ、SNAP4、CBCL、AQ、SRS、SAICA
- 每筆資料以 `famid` + `record_date` + `role` 作為唯一鍵，已存在的資料不覆寫
- 欄位值為 `999` 時存入 `null`，跳過範圍驗證
- 欄位值超出有效範圍時，整筆資料不匯入，另存至 `errors_YYYYMMDD_HHMMSS.xlsx`
