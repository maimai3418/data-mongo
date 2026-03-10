from dotenv import load_dotenv
from src.reader import read_xlsx
from src.transformer import split_by_collection
from src.importer import get_db, upsert_many
from src.logger import log_summary
from src.error_writer import write_error_xlsx

load_dotenv()

FILEPATH = "import_data.xlsx" 

def main():
    print("reading file...")
    df = read_xlsx(FILEPATH)
    print(f"total rows: {len(df)}")

    print("splitting columns...")
    collections, error_rows = split_by_collection(df)

    print("writing errors to xlsx...")
    write_error_xlsx(error_rows)

    print("writing to MongoDB...")
    db = get_db()
    total_rows = len(df)
    for col_name, docs in collections.items():
        error_count = len([r for r in error_rows if r.get("collection") == col_name])
        assert len(docs) + error_count == total_rows, f"[{col_name}] count mismatch!"
        log_summary(col_name, len(docs))
        upsert_many(db, col_name, docs)

    print("done")

if __name__ == "__main__":
    main()