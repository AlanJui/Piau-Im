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
from mod_excel_access import set_range_format

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
    sheet = wb.sheets['漢字注音']  # 選擇【漢字注音】工作表

    # 從 env 工作表中獲取每頁總列數和每列總字數
    env_sheet = wb.sheets['env']
    total_lines = int(env_sheet.range('每頁總列數').value)
    chars_per_row = int(env_sheet.range('每列總字數').value)

    # 設定起始及結束的【列】位址
    ROWS_PER_LINE = 4
    start_row = 5
    end_row = start_row + (total_lines * ROWS_PER_LINE)

    # 設定起始及結束的【欄】位址
    start_col = 4  # D 欄
    end_col = start_col + chars_per_row - 1  # 因為欄位是從 1 開始計數

    # for row in range(start_row, end_row + 1, ROWS_PER_LINE):
    # 清除內容並設置格式
    row = start_row
    for line in range(1, total_lines + 1):
        # 判斷是否已經超過結束列位址，若是則跳出迴圈
        if row > end_row: break
        # 顯示目前處理【狀態】
        print(f'重置 {line} 行：【漢字】儲存格位於【 {row} 列 】。')

        # 人工標音
        range_人工標音 = sheet.range((row - 2, start_col), (row - 2, end_col))
        range_人工標音.value = None
        set_range_format(range_人工標音,
                         font_name='Arial',
                         font_size=24,
                         font_color=0xFF0000,   # 紅色
                         fill_color=0xFFFFCC)

        # 台語音標
        range_台語音標 = sheet.range((row - 1, start_col), (row - 1, end_col))
        range_台語音標.value = None
        set_range_format(range_台語音標,
                         font_name='Sitka Text Semibold',
                         font_size=24,
                         font_color=0xFF9933)  # 橙色

        # 漢字
        range_漢字 = sheet.range((row, start_col), (row, end_col))
        range_漢字.value = None
        set_range_format(range_漢字,
                         font_name='吳守禮細明台語注音',
                         font_size=48,
                         font_color=0x000000)  # 黑色

        # 漢字標音
        range_漢字標音 = sheet.range((row + 1, start_col), (row + 1, end_col))
        range_漢字標音.value = None
        set_range_format(range_漢字標音,
                         font_name='芫荽 0.94',
                         font_size=26,
                         font_color=0x009900)  # 綠色

        # 準備處理下一【行】
        row += ROWS_PER_LINE
    # 返回【作業正常結束代碼】
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
        print(f"程式發生異常問題: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    # =========================================================================
    # 結束作業
    # =========================================================================
    print("程式執行完畢！")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
