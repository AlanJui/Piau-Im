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

# 載入自訂模組/函式
from a701_作業中活頁檔填入漢字 import process as fill_hanji_in_cells

# 載入自訂模組/函式
from mod_excel_access import (
    get_han_ji_khoo,
    get_total_rows_in_sheet,
    maintain_han_ji_koo,
)
from mod_file_access import save_as_new_file
from mod_標音 import is_punctuation  # 是否為標點符號
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
def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im'):
    """
    將 Excel 工作表中指定區域的漢字取出，儲存為一個純文字檔。
    """
    # 取得【漢字庫】工作表物件
    han_ji_koo_sheet = get_han_ji_khoo(wb)

    # 選擇工作表
    sheet = wb.sheets['漢字注音']
    sheet.activate()
    #--------------------------------------------------------------------------
    # 自【env】設定工作表，取得處理作業所需參數
    #--------------------------------------------------------------------------

    # 設定起始及結束的【列】位址（【第5列】、【第9列】、【第13列】等列）
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    line = 1

    # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start_col = 4
    end_col = start_col + CHARS_PER_ROW

    #--------------------------------------------------------------------------
    # 作業處理：逐列取出漢字，組合成純文字檔
    #--------------------------------------------------------------------------
    logging_process_step(f"開始【處理作業】...")
    han_ji_text = ""
    EOF = False

    # 逐列處理作業
    for row in range(start_row, end_row, ROWS_PER_LINE):
        # 若已到【結尾】或【超過總行數】，則跳出迴圈
        if EOF or line > TOTAL_LINES:
            break

        # 設定【作用儲存格】為列首
        Two_Empty_Cells = 0
        sheet.range((row, 1)).select()

        # 逐欄取出漢字處理
        for col in range(start_col, end_col):
            # 取得當前儲存格內含值
            cell_value = sheet.range((row, col)).value
            if cell_value == 'φ':       # 讀到【結尾標示】
                EOF = True
                msg = "【文字終結】"
            elif cell_value == '\n':    # 讀到【換行標示】
                han_ji_text += '\n'
                msg = "【換行】"
            elif cell_value == None:    # 讀到【空白】
                if Two_Empty_Cells == 0:
                    Two_Empty_Cells += 1
                elif Two_Empty_Cells == 1:
                    EOF = True
                msg = "【缺空】"    # 表【儲存格】未填入任何字/符，不同於【空白】字元
            else:                       # 讀到：漢字或標點符號
                han_ji_text += cell_value
                msg = cell_value
                if not is_punctuation(cell_value):
                    han_ji = cell_value
                    tai_gi = sheet.range((row - 1, col)).value
                    maintain_han_ji_koo(sheet=han_ji_koo_sheet,
                                        han_ji=han_ji, tai_gi=tai_gi,
                                        show_msg=False)

            # 顯示處理進度
            col_name = xw.utils.col_name(col)   # 取得欄位名稱
            print(f"({row}, {col_name}) = {msg}")

            # 若讀到【換行】或【文字終結】，跳出逐欄取字迴圈
            if msg == "【換行】" or EOF:
                break

        # 每當處理一行 15 個漢字後，亦換到下一行
        print("\n")
        line += 1
        row += 4

    # 作業結束前處理
    logging_process_step(f"完成【作業程序】：查找漢字注音...")
    return EXIT_CODE_SUCCESS


def update_han_ji_piau_im(wb, han_ji_koo_sheet_name='漢字庫', han_ji_piau_im_sheet_name='漢字注音'):
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【漢字庫】中的【校正】欄位進行更新。
    wb: Excel 活頁簿物件
    han_ji_koo_sheet_name: 【漢字庫】工作表名稱
    han_ji_zhu_yin_sheet_name: 【漢字注音】工作表名稱
    """
    # 取得工作表
    han_ji_koo_sheet = wb.sheets[han_ji_koo_sheet_name]
    han_ji_piau_im_sheet = wb.sheets[han_ji_piau_im_sheet_name]

    # 取得【漢字庫】表格範圍的所有資料
    data = han_ji_koo_sheet.range("A2").expand("table").value

    if data is None:
        print("【漢字庫】工作表無資料")
        return EXIT_CODE_INVALID_INPUT

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 將資料轉為字典格式，key: 漢字, value: (台語音標, 校正, 次數)
    han_ji_dict = {}
    for row in data:
        han_ji = row[0] or ""
        tai_gi_im_piau = row[1] or ""
        total_count = int(row[2]) if len(row) > 2 and isinstance(row[2], (int, float)) else 0
        corrected_tai_gi = row[3] if len(row) > 3 else ""  # 若無 D 欄資料則設為空字串

        if corrected_tai_gi and (corrected_tai_gi != tai_gi_im_piau):
            han_ji_dict[han_ji] = (tai_gi_im_piau, corrected_tai_gi, total_count)

    # 若無需更新的資料，結束函數
    if not han_ji_dict:
        print("【漢字庫】工作表中，【校正音標】欄，均未填入需更新之台語音標！")
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
                _, corrected_tai_gi, total_count = han_ji_dict[han_ji]
                tai_gi_cell = han_ji_piau_im_sheet.range((row - 1, col))
                original_tai_gi = tai_gi_cell.value or ""

                # 更新多次，直到總數用完
                if corrected_tai_gi != original_tai_gi and total_count > 0:
                    tai_gi_cell.value = corrected_tai_gi  # 更新儲存格
                    han_ji_cell.color = (255, 255, 0)       # 將底色設為【黄色】
                    han_ji_cell.font.color = (255, 0, 0)    # 將文字顏色設為【紅色】

                    msg = f"({row}, {xw.utils.col_name(col)}) = {han_ji}，台語音標由【{original_tai_gi}】改為【{corrected_tai_gi}】"
                    print(msg)
                    total_count -= 1  # 減少剩餘更新次數

                    # 更新完畢後，減少【漢字庫】的總數
                    han_ji_koo_sheet.range(f"C{row + 1}").value = total_count
                    if total_count == 0:
                        print(f"漢字【{han_ji}】的更新次數已用完")

    print("【漢字注音】表的台語音標更新作業已完成")

    # 作業結束前處理
    logging_process_step(f"完成【作業程序】：更新漢字標音...")
    return EXIT_CODE_SUCCESS


def process(wb):
    # return_code = ca_han_ji_thak_im(wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     logging_process_step("處理作業失敗，過程中出錯！")
    #     return return_code

    return_code = update_han_ji_piau_im(wb)
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