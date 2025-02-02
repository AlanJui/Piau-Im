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
from mod_標音 import convert_tl_to_tlpa

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_PROCESS_FAILURE = 3
EXIT_CODE_UNKNOWN_ERROR = 99

# =========================================================================
# 功能 5：依據工作表之資料，建置【漢字庫】資料表
# =========================================================================
def rebuild_database_from_excel(wb, sheet_name="漢字庫"):
    """
    依據 Excel 工作表的資料，重建 SQLite 資料庫的【漢字庫】資料表。
    - 刪除舊的【漢字庫】資料表。
    - 根據 Excel 工作表的資料重建資料表。
    - 轉換拼音 TL → TLPA。
    - 確保【識別號】為 PRIMARY KEY AUTOINCREMENT。
    - 建立 UNIQUE INDEX (漢字, 台羅音標) 避免重複。

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

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # **1️⃣ 刪除現有 `漢字庫` 資料表**
        cursor.execute("DROP TABLE IF EXISTS 漢字庫")

        # **2️⃣ 重新建立 `漢字庫` 資料表**
        cursor.execute("""
        CREATE TABLE 漢字庫 (
            識別號 INTEGER PRIMARY KEY AUTOINCREMENT,
            漢字 TEXT NOT NULL,
            台羅音標 TEXT NOT NULL,
            常用度 REAL DEFAULT 0.1,
            摘要說明 TEXT DEFAULT 'NA',
            更新時間 TEXT DEFAULT (DATETIME('now', 'localtime')) NOT NULL
        );
        """)

        # **3️⃣ 讀取 Excel 工作表資料**
        for idx, row_data in enumerate(data, start=2):  # Excel A2 起始，Python Index 2
            han_ji = row_data[1]  # B 欄: 漢字
            tai_lo_im_piau = row_data[2]  # C 欄: 台羅音標
            siong_iong_too = row_data[3] if isinstance(row_data[3], (int, float)) else 0.1  # D 欄: 常用度
            summary = row_data[4] if isinstance(row_data[4], str) else "NA"  # E 欄: 摘要說明
            updated_time = row_data[5] if isinstance(row_data[5], str) else datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # **Console Debug 訊息**
            print(f"📌 正在處理第 {idx-1} 筆資料 (Excel 第 {idx} 列): 漢字='{han_ji}', 台羅音標='{tai_lo_im_piau}', 更新時間='{updated_time}'")

            # **確保 `漢字` 和 `台羅音標` 務必要有資料**
            if not han_ji or not tai_lo_im_piau:
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列：缺【漢字】或【台羅音標】")
                # **將錯誤記錄寫入 `error_log.txt`**
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                continue  # 跳過無效資料

            # **檢查 `台羅音標` 是否為有效字串**
            if not han_ji or not isinstance(tai_lo_im_piau, str) or not tai_lo_im_piau.strip():
                print(f"⚠️ 跳過無效資料: Excel 第 {idx} 列 (台羅音標格式錯誤)")
                # **將錯誤記錄寫入 `error_log.txt`**
                with open("error_log.txt", "a", encoding="utf-8") as log_file:
                    log_file.write(f"❌ 無效資料（Excel 第 {idx} 列）: {row_data}\n")
                continue  # **跳過此筆錯誤資料**

            # **轉換台羅拼音（TL）→ 台語音標（TLPA）**
            tlpa_pinyin = convert_tl_to_tlpa(tai_lo_im_piau)

            cursor.execute("""
                INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 摘要說明, 更新時間)
                VALUES (?, ?, ?, ?, ?);
            """, (han_ji, tlpa_pinyin, siong_iong_too, summary, updated_time))

        # **4️⃣ 建立 `UNIQUE INDEX` 確保無重複**
        cursor.execute("CREATE UNIQUE INDEX idx_漢字_台羅音標 ON 漢字庫 (漢字, 台羅音標);")

        conn.commit()
        print("✅ `漢字庫` 資料表已成功重建！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 重建 `漢字庫` 失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        conn.close()

# =========================================================================
# 主程式執行
# =========================================================================
def main():
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "4"

    wb = xw.apps.active.books.active

    if mode == "4":
        return rebuild_database_from_excel(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (4)")
        return EXIT_CODE_INVALID_INPUT

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)