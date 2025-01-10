import logging
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_file_access import get_named_value, save_as_new_file
from p730_Tng_Sing_Bang_Iah_R1 import tng_sing_bang_iah

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
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def logging_process_step(msg):
    print(msg)
    logging.info(msg)

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


def process(wb):
    # (1) 指定作業使用：【漢字注音】工作表
    sheet = wb.sheets['漢字注音']   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    # (2) 將【漢字注音】工作表中的標音漢字，轉成 HTML 網頁檔案。
    result_code = tng_sing_bang_iah(
        wb=wb,
        sheet_name='漢字注音',
        cell='V3',
        page_type='含頁頭'
    )
    if result_code != EXIT_CODE_SUCCESS:
        logging.error("標音漢字轉換為 HTML 網頁檔案失敗！")
        return result_code

    # (3) 依 env 工作表之設定，將檔案儲存至指定目錄。
    file_path = save_as_new_file(wb=wb)
    if not file_path:
        logging.error("儲存檔案失敗！")
        return EXIT_CODE_PROCESS_FAILURE
    else:
        logging_process_step(f"儲存檔案至路徑：{file_path}")
        # 作業正常結束
        return EXIT_CODE_SUCCESS


def main():
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    print(f"專案根目錄為: {project_root}")
    logging.info(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 嘗試獲取當前作用中的 Excel 工作簿
    # =========================================================================
    wb = None
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        print("無法找到作用中的 Excel 工作簿")
        return EXIT_CODE_NO_FILE

    if not wb:
        print("無法作業，原因可能為：(1) 未指定輸入檔案；(2) 未找到作用中的 Excel 工作簿！")
        logging.error("無法作業，未指定輸入檔案或 Excel 無效。")
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 處理作業
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging.error("處理作業失敗，過程中出錯！")
            return result_code

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        logging.error(f"執行過程中發生未知錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            logging_process_step(f"製作【漢字標音】網頁己完成！")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("作業正常結束！")
    else:
        print(f"作業異常終結，異常碼為: {exit_code}")
    sys.exit(exit_code)
