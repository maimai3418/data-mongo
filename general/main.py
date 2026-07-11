import os
import sys

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from dotenv import load_dotenv
from src.reader import read_xlsx
from src.transformer import split_by_collection
from src.importer import get_db, upsert_many
from src.logger import log_summary
from src.error_writer import write_error_xlsx
from src.conflict_writer import write_conflict_xlsx
from src.import_logger import write_import_log
from src.utils.select_collections import select_collections
from src.utils.input_project_code import input_project_code
from src.utils.wait_and_retry import wait_and_retry

load_dotenv()

FILEPATH = "import_data.xlsx"

def main():
    # 輸入計畫代碼（選填）
    project_code = input_project_code()

    # 選擇要上傳的量表
    selected = select_collections()
    if selected is None:
        return
    print("reading file...")

    # 讀取 Excel 檔案
    df = read_xlsx(FILEPATH)
    print(f"total rows: {len(df)}")

    # print("splitting columns...")
    collections, error_rows, skipped_rows = split_by_collection(df, selected)

    # 記錄錯誤資訊到 errors.xlsx
    print("writing errors to xlsx...")
    wait_and_retry(lambda: write_error_xlsx(error_rows), "errors.xlsx")

    # 將資料寫入 MongoDB
    print("writing to MongoDB...")
    db = get_db()
    total_rows = len(df)
    log_entries = []
    all_conflict_rows = []

    for col_name, docs in collections.items():
        error_count = len([r for r in error_rows if r.get("collection") == col_name])
        skipped_count = len([r for r in skipped_rows if r.get("collection") == col_name])
        assert len(docs) + error_count + skipped_count == total_rows, f"[{col_name}] count mismatch!"
        log_summary(col_name, len(docs), skipped_count, error_count)
        if project_code:
            for doc in docs:
                doc["research_project_code"] = project_code
        result = upsert_many(db, col_name, docs)
        inserted, skipped_dup, conflict_docs, conflict_rows = (
            result if result else (0, 0, 0, [])
        )
        all_conflict_rows.extend(conflict_rows)
        entry = {
            "collection": col_name,
            "total": total_rows,
            "success": len(docs),
            "insert": inserted,
            "skipped_dup": skipped_dup,
            "value_conflict": conflict_docs,
            "errors": error_count,
            "skipped": skipped_count,
        }
        if project_code:
            entry["research_project_code"] = project_code
        log_entries.append(entry)

    # 匯出 value_conflict 衝突報告（key 相同但欄位值不同，未覆蓋 DB）
    if all_conflict_rows:
        print("writing conflict report...")
        wait_and_retry(lambda: write_conflict_xlsx(all_conflict_rows), "conflicts.xlsx")

    # 記錄本次匯入的統計資訊到 import_log.xlsx
    print("writing import log...")
    wait_and_retry(lambda: write_import_log(log_entries), "import_log.xlsx")

    print("done")

if __name__ == "__main__":
    main()