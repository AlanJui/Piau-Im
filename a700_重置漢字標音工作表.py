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
from mod_excel_access import reset_han_ji_cells
from mod_file_access import save_as_new_file

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
    #----------------------------------------------------------------------
    # 將儲存格內的舊資料清除
    #----------------------------------------------------------------------
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

    #--------------------------------------------------------------------------
    # 作業結尾處理
    #--------------------------------------------------------------------------
    file_path = save_as_new_file(wb=wb)
    if not file_path:
        logging.error("儲存檔案失敗！")
        return EXIT_CODE_PROCESS_FAILURE    # 作業異當終止：無法儲存檔案
    else:
        logging_process_step(f"儲存檔案至路徑：{file_path}")
        return EXIT_CODE_SUCCESS    # 作業正常結束

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
            logging.info("a700_重罝漢字標音工作表.py 程式已執行完畢！")

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
