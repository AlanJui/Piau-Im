# WSL2: /mnt/c/Users/AlanJui/AppData/Roaming/Rime
# C:\Users\AlanJui\AppData\Roaming\Rime
# WSL2: /mnt/z/home/alanjui/workspace/rime/rime-tlpa
# Z:\home\alanjui\workspace\rime\rime-tlpa
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
# 功能 4：將資料庫之【漢字庫】資料表，備份至 Excel 工作表
# =========================================================================
def export_database_to_excel(wb, sheet_name="漢字庫"):
    """
    將 `漢字庫` 資料表的資料寫入 Excel 的【漢字庫】工作表。

    :param wb: Excel 活頁簿物件
    :return: EXIT_CODE_SUCCESS or EXIT_CODE_FAILURE
    """
    try:
        ensure_sheet_exists(wb, sheet_name)
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"⚠️ 無法找到工作表: {sheet_name}")
        return EXIT_CODE_FAILURE

    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # 讀取資料庫內容
        cursor.execute("SELECT 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫;")
        rows = cursor.fetchall()

        # 清空舊內容
        sheet.clear()

        # 寫入標題列
        sheet.range("A1").value = ["識別號", "漢字", "台羅音標", "常用度", "摘要說明" "更新時間"]

        # 寫入資料
        sheet.range("A2").value = rows

        print("✅ 資料成功匯出至 Excel！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 匯出資料失敗: {e}")
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
        return export_database_to_excel(wb)
    else:
        print("❌ 錯誤：請輸入有效模式 (4)")
        return EXIT_CODE_INVALID_INPUT

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)