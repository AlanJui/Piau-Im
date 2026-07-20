# =========================================================================
# a850_使用Excel重建漢字庫.py
#
# 功能說明：
# 依據作用中 Excel 活頁簿之【漢字庫】工作表，重建【Ho_Lok_Ue.db】資料庫的【漢字庫】資料表。
# 與 a840_匯出漢字庫至Exccel.py 為【備份／還原】配對工具。
#
# 【漢字庫】工作表欄位結構（須與 a840 匯出格式一致）：
#   A 識別號、B 漢字、C 台羅音標、D 常用度、
#   E 摘要說明、F 更新時間、G 最近揀用時間
#
# 注意：
# - 【台羅音標】原樣寫回，不做任何拼音轉換（避免泉／漳腔拼寫被正規化而合併）。
# - 保留 Excel A 欄之【識別號】，令還原後之資料庫與備份前一致。
# - 本程式會 DROP 既有【漢字庫】資料表後重建，執行前請確認來源無誤。
#
# 用法：
#   1. 開啟由 a840 匯出之 Excel 檔（如「河洛話漢字庫.xlsx」），並使之處於作用中；
#   2. python a850_使用Excel重建漢字庫.py
#   3. 可選參數：工作表名稱（預設「漢字庫」）
# =========================================================================

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from datetime import datetime

import xlwings as xw
from dotenv import load_dotenv

from mod_excel_access import ensure_sheet_exists

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_PROCESS_FAILURE = 3
EXIT_CODE_UNKNOWN_ERROR = 99

TOTAL_COLS = 7  # A～G
PROGRESS_EVERY = 2000  # 每處理 N 筆印一次進度


def normalize_time(value):
    """將 Excel 儲存格之時間值正規化為 'YYYY-MM-DD HH:MM:SS' 字串；無資料回傳 None。"""
    if isinstance(value, datetime):
        # Excel 若將時間字串自動轉成【日期】值，xlwings 讀回為 datetime 物件
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, str) and value.strip():
        return value.strip()
    return None


# =========================================================================
# 功能：依據工作表之資料，建置【漢字庫】資料表
# =========================================================================
def rebuild_database_from_excel(wb, sheet_name="漢字庫"):
    """
    依據 Excel 工作表的資料，重建 SQLite 資料庫的【漢字庫】資料表。
    - 刪除舊的【漢字庫】資料表。
    - 根據 Excel 工作表的資料重建資料表（含【最近揀用時間】）。
    - 【台羅音標】原樣寫回，不做任何拼音轉換。
    - 保留 Excel A 欄之【識別號】。
    - 建立 UNIQUE INDEX (漢字, 台羅音標) 與查音用複合索引。

    :param wb: Excel 活頁簿物件
    :param sheet_name: 工作表名稱，預設為 "漢字庫"
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    try:
        ensure_sheet_exists(wb, sheet_name)
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
    if last_row < 2:
        print("⚠️ 工作表無資料可供重建。")
        return EXIT_CODE_FAILURE
    data = sheet.range((2, 1), (last_row, TOTAL_COLS)).value

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # 1️⃣ 刪除現有【漢字庫】資料表
        cursor.execute("DROP TABLE IF EXISTS 漢字庫")

        # 2️⃣ 重新建立【漢字庫】資料表
        # 【最近揀用時間】：由人工校正程式回寫，查音時於常用度相同之讀音間排定優先順序
        cursor.execute(
            """
            CREATE TABLE 漢字庫 (
                識別號 INTEGER PRIMARY KEY AUTOINCREMENT,
                漢字 TEXT NOT NULL,
                台羅音標 TEXT NOT NULL,
                常用度 REAL DEFAULT 0.1,
                摘要說明 TEXT DEFAULT 'NA',
                更新時間 TEXT DEFAULT (DATETIME('now', 'localtime')) NOT NULL,
                最近揀用時間 TEXT
            );
            """
        )

        # 3️⃣ 整理 Excel 列為待插入之參數清單
        records = []
        skipped = 0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for idx, row_data in enumerate(data, start=2):  # Excel 列號（含標題列，資料自第 2 列起）
            id_no = int(row_data[0]) if isinstance(row_data[0], (int, float)) else None
            han_ji = row_data[1]
            tai_lo_im_piau = row_data[2]
            siong_iong_too = row_data[3] if isinstance(row_data[3], (int, float)) else 0.1
            summary = row_data[4] if isinstance(row_data[4], str) and row_data[4].strip() else "NA"
            updated_time = normalize_time(row_data[5]) or now_str
            last_pick_time = normalize_time(row_data[6])  # 允許 NULL

            if id_no is None:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列：缺【識別號】")
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列，缺識別號）: {row_data}\n")
                skipped += 1
                continue

            if not han_ji or tai_lo_im_piau is None:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列：缺【漢字】或【台羅音標】")
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                skipped += 1
                continue

            # 台羅音標：強制為字串並去除首尾空白（防 Excel 誤轉型後仍可還原）
            tai_lo_str = str(tai_lo_im_piau).strip()
            if not tai_lo_str:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列（台羅音標空白）")
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列，台羅音標空白）: {row_data}\n")
                skipped += 1
                continue

            records.append(
                (id_no, str(han_ji), tai_lo_str, siong_iong_too, summary, updated_time, last_pick_time)
            )

            if len(records) % PROGRESS_EVERY == 0:
                print(f"… 已整理 {len(records)} 筆 …")

        if not records:
            print("❌ 無有效資料可寫入，中止重建。")
            return EXIT_CODE_FAILURE

        # 4️⃣ 批次寫入
        cursor.executemany(
            """
            INSERT INTO 漢字庫 (識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間)
            VALUES (?, ?, ?, ?, ?, ?, ?);
            """,
            records,
        )

        # 5️⃣ 建立索引
        cursor.execute("CREATE UNIQUE INDEX idx_漢字_台羅音標 ON 漢字庫 (漢字, 台羅音標);")
        # 查音用複合索引（與 mod_程式.py / mod_ca_ji_tian.py 之查音排序規則一致）
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_漢字庫_查音 ON 漢字庫 (漢字, 常用度 DESC, 最近揀用時間 DESC);"
        )

        conn.commit()
        print(f"✅ 【漢字庫】資料表已成功重建！（寫入 {len(records)} 筆，跳過 {skipped} 筆）")
        logging.info("漢字庫重建完成：寫入 %s 筆，跳過 %s 筆", len(records), skipped)
        return EXIT_CODE_SUCCESS

    except Exception as e:
        conn.rollback()
        print(f"❌ 重建【漢字庫】失敗: {e}")
        logging.error("重建漢字庫失敗: %s", e)
        return EXIT_CODE_FAILURE

    finally:
        conn.close()


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    sheet_name = sys.argv[1] if len(sys.argv) > 1 else "漢字庫"

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"❌ 無法取得作用中的 Excel 活頁簿: {e}")
        return EXIT_CODE_FAILURE

    if not wb:
        print("❌ 無法作業，因未有任何 Excel 檔案已開啟。")
        return EXIT_CODE_FAILURE

    print(f"📌 來源活頁簿：{wb.name}／工作表：{sheet_name}")
    print(f"📌 目標資料庫：{DB_HO_LOK_UE}")
    return rebuild_database_from_excel(wb, sheet_name)


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
