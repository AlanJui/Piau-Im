
# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_file_access import (
    copy_excel_sheet,
    reset_han_ji_piau_im_cells,
    save_as_new_file,
)
from p704_漢字以十五音標注音 import han_ji_piau_im

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

# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    # (0) 取得專案根目錄。
    # 使用已打開且處於作用中的 Excel 工作簿
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        print(f"發生錯誤: {e}")
        print("無法找到作用中的 Excel 工作簿")
        sys.exit(2)

    # 獲取活頁簿的完整檔案路徑
    file_path = wb.fullname
    print(f"完整檔案路徑: {file_path}")

    # 獲取活頁簿的檔案名稱（不包括路徑）
    file_name = wb.name
    print(f"檔案名稱: {file_name}")

    # 顯示「已輸入之拼音字母及注音符號」
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    # (1) A720: 將 V3 儲存格內的漢字，逐個填入標音用方格。
    sheet = wb.sheets['漢字注音']
    sheet.activate()
    sheet.range('A1').select()

    # (2) 複製【漢字注音】工作表，並將【漢字注音】工作表已有漢字標清除（不含上列之【台語音標】）
    piau_im_huat = wb.names['標音方法'].refers_to_range.value

    copy_excel_sheet(wb, '漢字注音', piau_im_huat)
    reset_han_ji_piau_im_cells(wb, piau_im_huat)

    # 呼叫 han_ji_piau_im 函數，並傳入動態參數
    han_ji_piau_im(wb, sheet_name=piau_im_huat, cell='V3')

    # (3) A740: 將【漢字注音】工作表的內容，轉成 HTML 網頁檔案。
    # tng_sing_bang_iah(wb, '漢字注音', 'V3')

    # (4) A750: 將 Tai_Gi_Zu_Im_Bun.xlsx 檔案，依 env 工作表的設定，另存新檔到指定目錄。
    save_as_new_file(wb=wb)

    return EXIT_CODE_SUCCESS

# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        print(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("作業異常終止！")
            return result_code

    except Exception as e:
        print(f"作業過程發生未知的異常錯誤: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.info("a704_製作其它漢字標音.py 程式已執行完畢！")

    # =========================================================================
    # 結束作業
    # =========================================================================
    file_path = save_as_new_file(wb=wb, input_file_name='_working')
    if not file_path:
        logging.error("儲存檔案失敗！")
        return EXIT_CODE_PROCESS_FAILURE    # 作業異當終止：無法儲存檔案
    else:
        logging_process_step(f"儲存檔案至路徑：{file_path}")
        return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
