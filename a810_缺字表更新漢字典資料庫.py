# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime

import xlwings as xw
from dotenv import load_dotenv

from mod_excel_access import get_value_by_name
from mod_標音 import convert_tlpa_to_tl

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
# 功能 1：使用【缺字表】更新【漢字庫】資料表
# =========================================================================
def update_database_from_missing_characters(wb):
    """
    使用【缺字表】工作表的資料更新 SQLite 資料庫的【漢字庫】資料表。
    - 將【台語音標】轉換為【台羅音標】後寫入資料庫。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    sheet_name = "缺字表"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    # 讀取【語音類型】以便設定【常用度】
    gu_im_lui_hing = get_value_by_name(wb=wb, name="語音類型")
    # 確定 `常用度`（文讀音 0.8 / 白話音 0.6）
    siong_iong_too = 0.8 if gu_im_lui_hing == "文讀音" else 0.6

    # 讀取資料範圍
    data = sheet.range("A2").expand("table").value  # 讀取所有資料

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        for idx, row_data in enumerate(data, start=2):  # Excel A2 起始，Python Index 2
            han_ji = row_data[0]  # A 欄: 漢字
            tai_lo_im_piau = row_data[2]  # C 欄: 台語音標

            if not han_ji or not tai_lo_im_piau or tai_lo_im_piau == "N/A":
                continue  # 跳過無效資料

            # **轉換台語音標（TLPA）→ 台羅音標（TL）**
            tl_im_piau = convert_tlpa_to_tl(tai_lo_im_piau)

            # **在 INSERT 之前，顯示 Console 訊息**
            print(f"📌 寫入資料庫: 漢字='{han_ji}', 台語音標='{tai_lo_im_piau}', 轉換後台羅音標='{tl_im_piau}', Excel 第 {idx} 列")

            cursor.execute("""
                INSERT INTO 漢字庫 (漢字, 台羅音標, 常用度, 摘要說明, 更新時間)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(漢字, 台羅音標) DO UPDATE
                SET 更新時間=CURRENT_TIMESTAMP;
            """, (han_ji, tl_im_piau, siong_iong_too, "NA", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

        conn.commit()
        print("✅ 資料庫更新完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 資料庫更新失敗: {e}")
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
        mode = "1"

    wb = xw.apps.active.books.active

    if mode == "1":
        return update_database_from_missing_characters(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (1)")
        return EXIT_CODE_INVALID_INPUT

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)