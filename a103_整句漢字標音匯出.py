# ========================================================================
# 程式名稱：a103_整句漢字標音匯出.py
# 程式說明：根據【作用儲存格】位置，將整段句子的漢字和標音匯出顯示於 console。
# ========================================================================

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

#--------------------------------------------------------------------------
# 儲存格位置常數
#  - 每 1 【行】，內含 4 row ；第 1 行之 row no 為：3
#  - row 1: 人工標音儲存格 ===> row_no= 3,  7, 11, ...
#  - row 2: 台語音標儲存格 ===> row_no= 4,  8, 12, ...
#  - row 3: 漢字儲存格     ===> row_no= 5,  9, 13, ...
#  - row 4: 漢字標音儲存格 ===> row_no= 6, 10, 14, ...
#
# 依【作用儲存格】的 row no 求得：line_no = ((row_no - start_row_no) // rows_per_line) + 1
#
# 依【line_no】求得【基準列 row no】：base_row_no = start_row_no + ((line_no - 1) * rows_per_line)
#--------------------------------------------------------------------------
ROWS_PER_LINE = 4
START_ROW = 3  # 第 1 行的起始列號
START_COL = 4  # D 欄
END_COL = 18   # R 欄

TAI_GI_PIAU_IM_OFFSET = 1
HAN_JI_OFFSET = 2
HAN_JI_PIAU_IM_OFFSET = 3


# =========================================================================
# Local Function
# =========================================================================
def read_sentence_from_row(sheet, start_row_no, start_col, end_col, check_han_ji_row=None):
    """
    從指定的 row 開始讀取整段句子，直到遇到 '\n' 為止。
    如果一行讀完還沒遇到 '\n'，會繼續讀取下一行。

    Args:
        sheet: Excel 工作表
        start_row_no: 起始列號
        start_col: 起始欄號（D=4）
        end_col: 結束欄號（R=18）
        check_han_ji_row: 若非 None，則檢查此列號是否有換行符號（用於標音列）

    Returns:
        str: 讀取到的句子內容
    """
    content = ""
    current_row = start_row_no
    han_ji_row = check_han_ji_row

    while True:
        for col in range(start_col, end_col + 1):  # +1 以包含 R 欄（18）
            # 如果是標音列，先檢查對應的漢字列是否有換行符號
            if han_ji_row is not None:
                han_ji_value = sheet.range((han_ji_row, col)).value
                if han_ji_value == '\n' or han_ji_value == 'φ':
                    # 對應的漢字列有換行或結束符號，立即返回
                    return content.strip()

            # 讀取當前儲存格的值
            cell_value = sheet.range((current_row, col)).value

            if cell_value == '\n':
                # 遇到換行符號，立即返回已讀取的內容
                return content.strip()
            elif cell_value == 'φ':
                # 遇到結束符號，立即返回已讀取的內容
                return content.strip()
            elif cell_value and cell_value != '':
                content += cell_value + ' '

        # 這一行讀完了，但沒遇到 '\n'，繼續讀下一行
        current_row += ROWS_PER_LINE
        if han_ji_row is not None:
            han_ji_row += ROWS_PER_LINE

        # 安全檢查：避免無窮迴圈
        if current_row > 200:  # 假設不會超過 200 列
            break

    return content.strip()


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb):
    """
    根據作用儲存格位置，讀取並顯示整段句子的漢字和標音。
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業開始！---------->")

    # 選擇工作表
    sheet = wb.sheets['漢字注音']

    #--------------------------------------------------------------------------
    # （1）依【作用儲存格】所在處，求得 Excel 儲存格所在 row no
    #--------------------------------------------------------------------------
    active_cell = sheet.range(xw.apps.active.selection.address)
    current_row_no = active_cell.row

    print(f"作用儲存格位置：{active_cell.address}")
    print(f"作用儲存格列號：{current_row_no}")

    #--------------------------------------------------------------------------
    # （2）依 row no 求得：line_no = (row_no - start_row_no + 1) // 4
    #--------------------------------------------------------------------------
    line_no = ((current_row_no - START_ROW) // ROWS_PER_LINE) + 1
    base_row_no = START_ROW + ((line_no - 1) * ROWS_PER_LINE)
    print(f"作用儲存格行號：{line_no}")
    print(f"基準列號：{base_row_no}")

    #--------------------------------------------------------------------------
    # （3）依 line_no 求得【漢字】儲存格之 row_no
    #--------------------------------------------------------------------------
    han_ji_row_no = base_row_no + HAN_JI_OFFSET
    print(f"漢字列號：{han_ji_row_no}")

    #--------------------------------------------------------------------------
    # （4）依 line_no 求得【漢字標音】儲存格之 row_no
    #--------------------------------------------------------------------------
    han_ji_piau_im_row_no = base_row_no + HAN_JI_PIAU_IM_OFFSET
    print(f"標音列號：{han_ji_piau_im_row_no}")

    #--------------------------------------------------------------------------
    # （5）讀取整段句子的漢字和標音
    #--------------------------------------------------------------------------
    print("\n")

    # 讀取漢字
    han_ji_content = read_sentence_from_row(sheet, han_ji_row_no, START_COL, END_COL)
    print(han_ji_content)

    # 讀取標音（同時檢查對應的漢字列）
    piau_im_content = read_sentence_from_row(
        sheet, han_ji_piau_im_row_no, START_COL, END_COL,
        check_han_ji_row=han_ji_row_no
    )
    print(piau_im_content)

    print()

    # 作業結束前處理
    logging_process_step(f"完成【處理作業】...")
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
    # (2) 若無指定輸入檔案，則獲取當前作用中的 Excel 檔案。
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

    try:
        # =========================================================================
        # (3) 執行【處理作業】
        # =========================================================================
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
            logging.info("釋放 Excel 資源，處理完成。")

    # 結束作業
    logging.info("作業成功完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("作業正常結束！")
    else:
        print(f"作業異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)
