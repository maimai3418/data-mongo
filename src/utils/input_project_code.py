def input_project_code():
    """讓使用者輸入計畫代碼（選填），按 Enter 跳過。"""
    raw = input("請輸入計畫代碼（選填，按 Enter 跳過）：").strip()
    if raw:
        print(f"計畫代碼：{raw}")
        return raw
    print("未輸入計畫代碼，跳過。")
    return None
