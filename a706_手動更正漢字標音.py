# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from pathlib import Path
from typing import Callable

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from a701_作業中活頁檔填入漢字 import process as fill_hanji_in_cells

# 載入自訂模組/函式
# from p709_reset_han_ji_cells import reset_han_ji_cells
from a702_查找及填入漢字標音 import reset_han_ji_cells

# 載入自訂模組/函式
from mod_excel_access import get_han_ji_khoo, get_value_by_name, maintain_han_ji_koo
from mod_file_access import load_module_function, save_as_new_file
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉台語音標

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
# 本程式主要處理作業程序
# =========================================================================
def update_han_ji_piau_im(wb):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。
    """
    # 建置 PiauIm 物件，供作漢字拼音轉換作業
    han_ji_khoo_field = '漢字庫'
    han_ji_khoo_name = get_value_by_name(wb=wb, name=han_ji_khoo_field) # 取得【漢字庫】名稱：河洛話、廣韻
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)            # 指定漢字自動查找使用的【漢字庫】
    ue_im_lui_piat = get_value_by_name(wb=wb, name='語音類型')
    piau_im_huat = get_value_by_name(wb=wb, name='標音方法')    # 指定【台語音標】轉換成【漢字標音】的方法

    # 取得工作表
    han_ji_piau_im_sheet = wb.sheets['漢字注音']
    piau_im_sheet_name = '標音字庫'
    piau_im_ji_khoo_sheet = get_han_ji_khoo(wb=wb, sheet_name=piau_im_sheet_name)

    # 取得【標音字庫】表格範圍的所有資料
    data = piau_im_ji_khoo_sheet.range("A2").expand("table").value

    if data is None:
        print("【標音字庫】工作表無資料")
        return EXIT_CODE_INVALID_INPUT

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 將資料轉為字典格式，key: 漢字, value: (台語音標, 校正音標, 次數)
    han_ji_dict = {}
    for i, row in enumerate(data, start=2):
        han_ji = row[0] or ""
        tai_gi_im_piau = row[1] or ""
        total_count = int(row[2]) if len(row) > 2 and isinstance(row[2], (int, float)) else 0
        corrected_tai_gi = row[3] if len(row) > 3 else ""  # 若無 D 欄資料則設為空字串

        if corrected_tai_gi and corrected_tai_gi != tai_gi_im_piau:
            han_ji_dict[han_ji] = (tai_gi_im_piau, corrected_tai_gi, total_count, i)  # i 為資料列索引

    # 若無需更新的資料，結束函數
    if not han_ji_dict:
        print("【標音字庫】工作表中，【校正音標】欄，均未填入需更新之台語音標！")
        return EXIT_CODE_SUCCESS

    # 逐列處理【漢字注音】表
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)

    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    for row in range(start_row, end_row, ROWS_PER_LINE):
        for col in range(start_col, end_col):
            han_ji_cell = han_ji_piau_im_sheet.range((row, col))
            han_ji = han_ji_cell.value or ""

            if han_ji in han_ji_dict:
                original_tai_gi, corrected_tai_gi, total_count, row_index_in_ji_khoo = han_ji_dict[han_ji]
                tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
                original_tai_gi_in_sheet = tai_gi_cell.value or ""
                jin_kang_piau_im_cell = han_ji_piau_im_sheet.range((row + 2, col))
                jin_kang_piau_im = jin_kang_piau_im_cell.value or ""
                han_ji_piau_im_cell = han_ji_piau_im_sheet.range((row + 1, col))

                # 更新多次，直到總數用完
                if corrected_tai_gi != original_tai_gi_in_sheet and total_count > 0:
                    if jin_kang_piau_im:
                        # 若【人工標音】已有標音，則不進行更新
                        msg = f"({row}, {xw.utils.col_name(col)}) = {han_ji}，已有人工標音【{jin_kang_piau_im}】，故無更新"
                    else:
                        tai_gi_cell.value = corrected_tai_gi  # 更新【台語音標】儲存格
                        han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                        han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】

                        tai_gi_im_piau = corrected_tai_gi
                        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                            piau_im=piau_im,
                            piau_im_huat=piau_im_huat,
                            tai_gi_im_piau=tai_gi_im_piau
                        )
                        han_ji_piau_im_cell.value = han_ji_piau_im  # 更新【漢字標音】儲存格

                        # 更新【標音字庫】中原【台語音標】欄位內容
                        piau_im_ji_khoo_sheet.range(f"B{row_index_in_ji_khoo}").value = corrected_tai_gi
                        msg = f"({row}, {xw.utils.col_name(col)}) = {han_ji}，台語音標由【{original_tai_gi_in_sheet}】改為【{corrected_tai_gi}】/【{han_ji_piau_im}】"

                    print(msg)
                    total_count -= 1  # 減少剩餘更新次數

                    # 更新完畢後，減少【標音字庫】的總數
                    piau_im_ji_khoo_sheet.range(f"C{row_index_in_ji_khoo}").value = total_count
                    if total_count == 0:
                        print(f"漢字【{han_ji}】的更新次數已用完")

    print("【漢字注音】表的台語音標更新作業已完成")

    # 作業結束前處理
    logging_process_step(f"完成【作業程序】：更新漢字標音並同步【標音字庫】內容...")
    return EXIT_CODE_SUCCESS


def process(wb):
    # ---------------------------------------------------------------------
    # 依【標音字庫】的【校正音標】，更新【漢字注音】表中的【台語音標】
    # ---------------------------------------------------------------------
    return_code = update_han_ji_piau_im(wb=wb)
    if return_code != EXIT_CODE_SUCCESS:
        logging_process_step("處理作業失敗，過程中出錯！")
        return return_code

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
            # 是否關閉 Excel 視窗可根據需求決定
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
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
