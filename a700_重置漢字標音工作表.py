import logging
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from p709_reset_han_ji_cells import reset_han_ji_cells

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

# =========================================================================
# 定義 Exit Code
# =========================================================================
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 定義 Exit Code
# =========================================================================
def logging_process_step(msg):
    print(msg)
    logging.info(msg)

# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb):
    #--------------------------------------------------------------------------
    # 將儲存格內的舊資料清除
    #--------------------------------------------------------------------------
    sheet = wb.sheets['漢字注音']   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    total_rows = wb.names['每頁總列數'].refers_to_range.value
    cells_per_row = 4
    end_of_rows = int((total_rows * cells_per_row ) + 2)
    cells_range = f'D3:R{end_of_rows}'

    sheet.range(cells_range).clear_contents()     # 清除 C3:R{end_of_row} 範圍的內容

    # 獲取 V3 儲存格的合併範圍
    merged_range = sheet.range('V3').merge_area
    # 清空合併儲存格的內容
    merged_range.clear_contents()

    #--------------------------------------------------------------------------
    # 將待注音的【漢字儲存格】，文字顏色重設為黑色（自動 RGB: 0, 0, 0）；填漢顏色重設為無填滿
    #--------------------------------------------------------------------------
    logging_process_step(f"開始【漢字注音】工作表的清空、重置！")
    if reset_han_ji_cells(wb) == EXIT_CODE_SUCCESS:
        logging_process_step(f"完成【漢字注音】工作表的清空、重置！")
    else:
        logging_process_step(f"【漢字注音】工作表的清空、重置失敗！")
        return EXIT_CODE_PROCESS_FAILURE

    return EXIT_CODE_SUCCESS


# =========================================================================
# 程式主要作業流程
# =========================================================================
def main():
    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案並另存新檔。
    # =========================================================================
    wb = None
    # 使用已打開且處於作用中的 Excel 工作簿
    try:
        # 嘗試獲取當前作用中的 Excel 工作簿
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_process_step(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    if not wb:
        logging_process_step("無法作業，因未無任何 Excel 檔案己開啟。")
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("處理作業失敗，過程中出錯！")
            return result_code

    except Exception as e:
        print(f"執行過程中發生未知錯誤: {e}")
        logging.error(f"執行過程中發生未知錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            wb.save()
            # 是否關閉 Excel 視窗可根據需求決定
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.info("釋放 Excel 資源，處理完成。")

    # =========================================================================
    # 結束作業
    # =========================================================================
    logging.info("作業成功完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("作業正常結束！")
    else:
        print(f"作業異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
