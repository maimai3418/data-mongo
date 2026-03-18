def wait_and_retry(action, description):
    while True:
        try:
            action()
            return
        except PermissionError:
            input(f"\n⚠️  無法寫入：{description} 可能已被開啟，請關閉後按 Enter 繼續...")