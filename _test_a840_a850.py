"""一次性測試：a840 匯出 → a850 重建（使用臨時活頁簿與臨時測試資料庫，不觸碰正式資料庫）"""
import contextlib
import importlib
import os
import shutil
import sqlite3

import xlwings as xw

a840 = importlib.import_module("a840_匯出漢字庫至Exccel")
a850 = importlib.import_module("a850_使用Excel重建漢字庫")

TEST_DB = "test_rebuild.db"

wb = xw.Book()  # 新開臨時活頁簿
try:
    # (1) a840：自正式資料庫匯出至臨時活頁簿
    rc = a840.export_database_to_excel(wb)
    print("a840 exit code:", rc)
    sheet = wb.sheets["漢字庫"]
    print("標題列:", sheet.range("A1:G1").value)
    print("第2列:", sheet.range("A2:G2").value)
    last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
    print("資料列數:", last_row - 1)

    # (2) a850：改用臨時測試資料庫重建（先放一張假的舊漢字庫表，驗證會被 DROP 重建）
    conn = sqlite3.connect(TEST_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS 漢字庫 (識別號 INTEGER, 漢字 TEXT)")
    conn.commit()
    conn.close()
    a850.DB_HO_LOK_UE = TEST_DB
    with open("_a850_test_out.txt", "w", encoding="utf-8") as fout, contextlib.redirect_stdout(fout):
        rc = a850.rebuild_database_from_excel(wb)
    print("a850 exit code:", rc)
    with open("_a850_test_out.txt", "r", encoding="utf-8") as fin:
        tail = fin.readlines()[-5:]
    print("a850 輸出尾段:", *tail, sep="\n  ")

    # (3) 驗證重建結果
    src = sqlite3.connect("Ho_Lok_Ue.db")
    dst = sqlite3.connect(TEST_DB)
    n_src = src.execute("SELECT COUNT(*) FROM 漢字庫").fetchone()[0]
    n_dst = dst.execute("SELECT COUNT(*) FROM 漢字庫").fetchone()[0]
    print(f"正式庫筆數: {n_src}，重建庫筆數: {n_dst}")
    print("重建庫結構:", [r[1] for r in dst.execute("PRAGMA table_info(漢字庫)")])
    print("重建庫索引:", [r[0] for r in dst.execute("SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='漢字庫'")])
    print("重建庫【洵】:")
    for row in dst.execute("SELECT 識別號, 漢字, 台羅音標, 常用度, 更新時間, 最近揀用時間 FROM 漢字庫 WHERE 漢字=?", ("洵",)):
        print("   ", row)
    # 逐筆比對兩庫內容是否完全一致
    q = "SELECT 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間 FROM 漢字庫 ORDER BY 識別號"
    identical = src.execute(q).fetchall() == dst.execute(q).fetchall()
    print("兩庫逐筆比對完全一致:", identical)
    src.close()
    dst.close()
finally:
    wb.close()  # 關閉臨時活頁簿（不儲存）
    if os.path.exists(TEST_DB):
        os.remove(TEST_DB)
        print("已刪除臨時測試資料庫。")
