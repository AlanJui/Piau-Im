"""
一次性資料修復腳本：將 Ho_Lok_Ue.db 中【河洛話】影子資料表的校正紀錄，合併回【漢字庫】資料表。

背景：
    mod_程式.py 曾誤將活頁簿命名範圍【漢字庫】之值（"河洛話"）當成資料表名稱，
    導致 a250/a260/a270 等人工校正程式，自 2026-01-11 起將校正結果寫入
    一張意外建立的【河洛話】影子資料表，而非查音時實際使用的【漢字庫】資料表。
    程式錯誤已於 2026-07-18 修正；本腳本負責把影子表累積的校正紀錄併回【漢字庫】。

合併規則：
    1. 剔除無效紀錄：【漢字】非漢字（如標點）、【台羅音標】含非 ASCII 字元（如帶調符之誤存資料）。
    2. 影子表內同（漢字, 台羅音標）有多筆者，取【最近揀用時間】（若無則【更新時間】）最新一筆。
    3. 併入【漢字庫】：
       - 已存在同（漢字, 台羅音標）者：更新
         * 常用度 = max(既有值, 影子表值)（只升不降，避免動到 1.0 之文白通用音）
         * 更新時間、最近揀用時間 = 取兩者較新者
       - 不存在者：新增，摘要說明註記「由河洛話影子表合併」。
    4. 正式執行（--apply）前，先將資料庫檔案備份為 Ho_Lok_Ue.backup-<時戳>.db；
       合併完成後，將【河洛話】資料表更名為【河洛話_已併入漢字庫_<日期>】留作稽核。

用法：
    python tools/merge_河洛話_to_漢字庫.py           # 試跑（dry-run）：僅列出將執行之動作，不寫入
    python tools/merge_河洛話_to_漢字庫.py --apply   # 正式執行合併
"""

import argparse
import shutil
import sqlite3
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / "Ho_Lok_Ue.db"
SOURCE_TABLE = "河洛話"
TARGET_TABLE = "漢字庫"


def is_han_ji(text: str) -> bool:
    """判斷字串是否為單一漢字（CJK 統一表意文字）。"""
    if not text or len(text) != 1:
        return False
    return unicodedata.category(text) == "Lo" and "CJK" in unicodedata.name(text, "")


def is_valid_im_piau(im_piau: str) -> bool:
    """台羅音標應為純 ASCII 之字母＋數字調號（帶調符者為誤存資料）。"""
    return bool(im_piau) and im_piau.isascii() and im_piau.isalnum()


def pick_time(row: dict) -> str:
    """取影子表紀錄之【揀用時間】：優先【最近揀用時間】，否則【更新時間】。"""
    return row["最近揀用時間"] or row["更新時間"] or ""


def main() -> int:
    parser = argparse.ArgumentParser(description="將【河洛話】影子資料表合併回【漢字庫】")
    parser.add_argument("--apply", action="store_true", help="正式執行合併（預設為試跑，不寫入）")
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"❌ 找不到資料庫檔案：{DB_PATH}")
        return 1

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    if SOURCE_TABLE not in tables:
        print(f"❌ 資料庫中無【{SOURCE_TABLE}】資料表（或已合併完成）。現有資料表：{tables}")
        return 1

    rows = [dict(r) for r in cur.execute(f"SELECT * FROM {SOURCE_TABLE} ORDER BY 識別號")]
    print(f"影子表【{SOURCE_TABLE}】共 {len(rows)} 筆紀錄。")

    # (1) 剔除無效紀錄
    skipped = []
    valid_rows = []
    for row in rows:
        if not is_han_ji(row["漢字"] or "") or not is_valid_im_piau(row["台羅音標"] or ""):
            skipped.append(row)
        else:
            valid_rows.append(row)

    # (2) 同（漢字, 台羅音標）去重，取揀用時間最新一筆
    merged: dict[tuple, dict] = {}
    for row in valid_rows:
        key = (row["漢字"], row["台羅音標"])
        if key not in merged or pick_time(row) > pick_time(merged[key]):
            merged[key] = row

    print(f"剔除無效紀錄 {len(skipped)} 筆；去重後待合併 {len(merged)} 組（漢字, 台羅音標）。")
    if skipped:
        print("\n=== 剔除之無效紀錄 ===")
        for row in skipped:
            print(f"  識別號 {row['識別號']}: 漢字={row['漢字']!r}, 台羅音標={row['台羅音標']!r}")

    # (3) 逐組比對【漢字庫】，決定動作
    updates = []  # (目標識別號, 新常用度, 新更新時間, 新揀用時間, 來源row, 既有row)
    inserts = []  # 來源row
    for (han_ji, im_piau), src in sorted(merged.items()):
        targets = [dict(r) for r in cur.execute(
            f"SELECT * FROM {TARGET_TABLE} WHERE 漢字 = ? AND 台羅音標 = ?", (han_ji, im_piau)
        )]
        src_time = pick_time(src)
        if targets:
            for tgt in targets:
                new_siong_iong = max(tgt["常用度"] or 0.0, src["常用度"] or 0.0)
                new_update = max(tgt["更新時間"] or "", src["更新時間"] or "")
                new_pick = max(tgt["最近揀用時間"] or "", src_time)
                updates.append((tgt["識別號"], new_siong_iong, new_update, new_pick, src, tgt))
        else:
            inserts.append(src)

    print(f"\n=== 更新既有紀錄：{len(updates)} 筆 ===")
    for tgt_id, new_sit, new_upd, new_pick, src, tgt in updates:
        changes = []
        if (tgt["常用度"] or 0.0) != new_sit:
            changes.append(f"常用度 {tgt['常用度']} → {new_sit}")
        if (tgt["最近揀用時間"] or "") != new_pick:
            changes.append(f"最近揀用時間 {tgt['最近揀用時間']} → {new_pick}")
        note = "、".join(changes) if changes else "（無實質變更）"
        print(f"  漢字庫#{tgt_id} {tgt['漢字']}/{tgt['台羅音標']}：{note}")

    print(f"\n=== 新增紀錄：{len(inserts)} 筆 ===")
    for src in inserts:
        print(f"  {src['漢字']}/{src['台羅音標']}：常用度={src['常用度']}, 揀用時間={pick_time(src)}")

    if not args.apply:
        print("\n【試跑模式】未寫入任何變更。確認無誤後，請加 --apply 參數正式執行。")
        conn.close()
        return 0

    # (4) 正式執行：先備份資料庫檔案
    conn.close()
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = DB_PATH.with_name(f"{DB_PATH.stem}.backup-{timestamp}{DB_PATH.suffix}")
    shutil.copy2(DB_PATH, backup_path)
    print(f"\n✅ 已備份資料庫：{backup_path.name}")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        for tgt_id, new_sit, new_upd, new_pick, _src, _tgt in updates:
            cur.execute(
                f"UPDATE {TARGET_TABLE} SET 常用度 = ?, 更新時間 = ?, 最近揀用時間 = ? WHERE 識別號 = ?",
                (new_sit, new_upd, new_pick or None, tgt_id),
            )
        for src in inserts:
            cur.execute(
                f"""INSERT INTO {TARGET_TABLE} (漢字, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                (
                    src["漢字"],
                    src["台羅音標"],
                    src["常用度"],
                    "由河洛話影子表合併",
                    src["更新時間"],
                    pick_time(src) or None,
                ),
            )
        archive_name = f"{SOURCE_TABLE}_已併入漢字庫_{datetime.now().strftime('%Y%m%d')}"
        cur.execute(f"ALTER TABLE {SOURCE_TABLE} RENAME TO {archive_name}")
        conn.commit()
        print(f"✅ 合併完成：更新 {len(updates)} 筆、新增 {len(inserts)} 筆。")
        print(f"✅ 影子表已更名為【{archive_name}】留作稽核，確認無誤後可自行刪除。")
    except Exception as e:
        conn.rollback()
        print(f"❌ 合併失敗，已回滾：{e}")
        return 1
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
