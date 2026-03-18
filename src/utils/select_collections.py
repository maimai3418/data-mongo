from config import COLLECTION_MAP

def select_collections():
    available = list(COLLECTION_MAP.keys())
    print("\n可用的量表：")
    for i, name in enumerate(available, 1):
        print(f"  {i}. {name}")
    print(f"  0. 全部 ({', '.join(available)})")
    print(f"  q. 取消")
    print()

    while True:
        raw = input("請選擇要上傳的量表（輸入編號，多個用逗號分隔，例如 1,3,5；0 代表全部；q 取消）：").strip()
        if raw.lower() == "q":
            print("已取消。")
            return None
        if raw == "0":
            return available
        try:
            indices = [int(x.strip()) for x in raw.split(",")]
            selected = [available[i - 1] for i in indices if 1 <= i <= len(available)]
            if selected:
                print(f"\n已選擇：{', '.join(selected)}")
                confirm = input("確認？(y/n)：").strip().lower()
                if confirm == "y":
                    return selected
        except (ValueError, IndexError):
            pass
        print("輸入有誤，請重新輸入。\n")