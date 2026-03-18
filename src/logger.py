def log_summary(collection_name: str, success: int, skipped: int, errors: int):
    print(f"[{collection_name}] 上傳: {success}, 跳過(空值): {skipped}, 錯誤: {errors}")