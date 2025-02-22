# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import (
    convert_to_excel_address,
    ensure_sheet_exists,
    excel_address_to_row_col,
    get_value_by_name,
)
from mod_標音 import convert_tl_to_tlpa, convert_tl_with_tiau_hu_to_tlpa

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()

# =========================================================================
# 程式區域函式
# =========================================================================
def update_excel_with_tai_gi(wb):
    """
    讀取 Excel 檔案，為 A 欄的每個漢字查詢台語音標與字義，並填入 C 欄（台語音標）與 F 欄（字義）。

    參數：
        file_path (str): Excel 檔案的路徑
    """
    try:
        sheet = wb.sheets["缺字表"]  # 選擇工作表
    except Exception as e:
        logging_exc_error(f"找不到名為「缺字表」的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    row = 2  # 從第 2 列開始（跳過標題列）
    while True:
        han_ji = sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束
            break

        # 取得【缺字表】中的【台語音標】（極有可能是【台羅拼音】且帶有調符與使用簡寫）與字義
        im_piau = sheet.range(f"C{row}").value
        tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)  # 將台語音標轉換為 TLPA+

        # 以經過轉換的【台語音標】更新【缺字表】的【校正音標】欄
        sheet.range(f"D{row}").value = tai_gi_im_piau

        print(f"{row-1}. (A{row}) 【{han_ji}】： 台語音標：{im_piau}, 校正音標：{tai_gi_im_piau}")

        row += 1  # 讀取下一行

    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
    program_name = current_file_path.stem

    # =========================================================================
    # 程式初始化
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # 開始執行程式
    # =========================================================================
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_exc_error(f"找不到作用中活頁簿檔", e)
        return EXIT_CODE_INVALID_INPUT

    status_code = update_excel_with_tai_gi(wb)
    if status_code != EXIT_CODE_SUCCESS:
        logging_process_step(f"程式執行失敗，錯誤代碼：{status_code}")
        return status_code

    return EXIT_CODE_SUCCESS

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
