#!/usr/bin/env python3
"""
famid → 計畫(project)對照。

定義：最外層資料夾 = 一個計畫。
遞迴遍歷 <tool_folder> 下所有 .xlsx / .xls（含多層子資料夾），每檔只取第一個 sheet，
讀取 famid 欄位，統計每個 famid 出現在哪幾個計畫、跨幾個計畫。

前提（已確認）：
    - 最外層資料夾就是計畫單位
    - famid 跨計畫共用同一套編碼（同一個 famid 在不同計畫指同一個案）

輸出：
    famid_project_map.csv   每列一個 famid：n_projects, projects, n_files, sources
    project_summary.csv     每個計畫的 famid 數、檔案數

用法：
    python famid_project_map.py <tool_folder>
    # <tool_folder> 要指到「各計畫資料夾的上一層」，parts[0] 才會抓到計畫名
"""

import sys
import os
import glob
import warnings
from datetime import datetime, timezone

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ----------------------------- CONFIG -----------------------------
FAMID_COL = "famid"
OUTPUT_DIR = "."
ZERO_PAD = 0   # >0 時把 famid 補零到固定寬度；0 表示不補
ROOT_LABEL = "(root)"   # 直接放在最外層、沒有子資料夾的檔案歸這類
# ------------------------------------------------------------------


def norm_famid_iter(values) -> set:
    """統一 famid：字串、去空白、去 .0 尾巴、去 NA/空，選擇性補零。"""
    out = set()
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if s == "" or s.lower() in ("nan", "none", "null"):
            continue
        if s.endswith(".0"):
            s = s[:-2]
        if ZERO_PAD > 0:
            s = s.zfill(ZERO_PAD)
        out.add(s)
    return out


def scan(folder: str):
    paths = []
    for ext in ("*.xlsx", "*.xls"):
        paths.extend(glob.glob(os.path.join(folder, "**", ext), recursive=True))
    paths = sorted(set(paths))
    if not paths:
        sys.exit(f"資料夾內(含子層)找不到 xlsx/xls：{folder}")

    famid_projects = {}   # famid -> {計畫}
    famid_sources = {}    # famid -> {相對路徑檔名}
    project_files = {}    # 計畫 -> 檔案數
    ok = warn = skip = 0

    for p in paths:
        rel = os.path.relpath(p, folder)
        parts = rel.split(os.sep)
        project = parts[0] if len(parts) > 1 else ROOT_LABEL
        try:
            df = pd.read_excel(p, sheet_name=0, dtype=str)
        except Exception as e:
            print(f"  [skip] 讀取失敗 {rel}: {e}")
            skip += 1
            continue
        if FAMID_COL not in df.columns:
            print(f"  [warn] {rel} 沒有欄位 '{FAMID_COL}'，跳過")
            warn += 1
            continue
        ids = norm_famid_iter(df[FAMID_COL].tolist())
        for fid in ids:
            famid_projects.setdefault(fid, set()).add(project)
            famid_sources.setdefault(fid, set()).add(rel)
        project_files[project] = project_files.get(project, 0) + 1
        ok += 1
        print(f"  [ok] {rel}: {len(ids)} 個 famid  (計畫: {project})")

    print(f"\n讀檔結果：成功 {ok}、無 famid 欄跳過 {warn}、讀取失敗 {skip}、總計 {len(paths)}")
    return famid_projects, famid_sources, project_files


def main():
    if len(sys.argv) < 2:
        sys.exit("用法: python famid_project_map.py <tool_folder>")
    folder = sys.argv[1]

    print(f"掃描資料夾: {folder}\n")
    famid_projects, famid_sources, project_files = scan(folder)

    all_famids = sorted(famid_projects)
    now = datetime.now(timezone.utc).isoformat()

    # --- 報表一：famid → 計畫對照 ---
    rows = []
    for fid in all_famids:
        projs = sorted(famid_projects[fid])
        srcs = sorted(famid_sources[fid])
        rows.append({
            "famid": fid,
            "n_projects": len(projs),
            "projects": ",".join(projs),
            "n_files": len(srcs),
            "sources": ",".join(srcs),
        })
    map_df = pd.DataFrame(rows).sort_values(
        ["n_projects", "famid"], ascending=[False, True]
    )
    map_path = os.path.join(OUTPUT_DIR, "famid_project_map.csv")
    map_df.to_csv(map_path, index=False, encoding="utf-8-sig")

    # --- 報表二：每個計畫摘要 ---
    proj_count = {}
    for fid, projs in famid_projects.items():
        for proj in projs:
            proj_count[proj] = proj_count.get(proj, 0) + 1
    summary_df = pd.DataFrame(
        [{"project": proj,
          "n_famid": proj_count[proj],
          "n_files": project_files.get(proj, 0)}
         for proj in sorted(proj_count)]
    ).sort_values("n_famid", ascending=False)
    summary_path = os.path.join(OUTPUT_DIR, "project_summary.csv")
    summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")

    # --- 終端摘要 ---
    cross = map_df[map_df["n_projects"] > 1]
    print("=" * 55)
    print(f"famid 總數（去重）: {len(all_famids)}")
    print(f"計畫數: {len(proj_count)}")
    print(f"跨 ≥2 計畫的 famid: {len(cross)}")
    print(f"famid→計畫對照: {map_path}")
    print(f"計畫摘要: {summary_path}")
    print(f"detected_at: {now}")


if __name__ == "__main__":
    main()