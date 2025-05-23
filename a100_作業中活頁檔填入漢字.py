# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import math
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from a200_查找及填入漢字標音 import ca_han_ji_thak_im

# 載入自訂模組
from mod_excel_access import (
    clear_han_ji_kap_piau_im,
    reset_cells_format_in_sheet,
    strip_cell,
)
from mod_file_access import save_as_new_file

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
# 用途：將漢字填入對應的儲存格
# 詳述：待加註讀音的漢字文置於 V3 儲存格。本程式將漢字逐字填入對應的儲存格：
# 【第一列】D5, E5, F5,... ,R5；
# 【第二列】D9, E9, F9,... ,R9；
# 【第三列】D13, E13, F13,... ,R13；
# 每個漢字佔一格，每格最多容納一個漢字。
# 漢字上方的儲存格（如：D4）為：【台語音標】欄，由【羅馬拼音字母】組成拼音。
# 漢字下方的儲存格（如：D6）為：【台語注音符號】欄，由【台語方音符號】組成注音。
# 漢字上上方的儲存格（如：D3）為：【人工標音】欄，可以只輸入【台語音標】；或
# 【台語音標】和【台語注音符號】皆輸入。

def fill_hanji_in_cells(wb, sheet_name='漢字注音', cell='V3'):
    # 選擇指定的工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    # 取得 V3 儲存格的字串
    v3_value = sheet.range(cell).value

    # 確認 V3 不為空
    if v3_value is None:
        logging_exc_error(msg="【待注音漢字】儲存格為空！", error=None)
        return EXIT_CODE_INVALID_INPUT

    # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start_col = 4
    # end_col = start_col + CHARS_PER_ROW - 1
    end_col = start_col + CHARS_PER_ROW

    # 設定起始及結束的【列】位址（【第5列】、【第9列】、【第13列】等列）
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value) if '每頁總列數' in [name.name for name in wb.names] else 120
    ROWS_PER_LINE = 4
    start_row = 5
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)

    # 逐字處理字串
    row = start_row
    col = start_col
    index = 0  # 用來追蹤目前處理到的字元位置
    total_length = len(v3_value)    # 計算字串的總長度

    while index < total_length:     # 使用 while 而非 for，確保處理完整個字串
        if row >= end_row: break
        # 設定當前作用儲存格，根據 `row` 和 `col` 動態選取
        sheet.range((row, 1)).select()

        # for col in range(start_col, end_col):  # 【D欄=4】到【R欄=18】
        while col < end_col:
            # 重置儲存格：文字顏色（黑色）及填滿色彩（無填滿）
            sheet.range((row-2, col), (row+1, col)).color = None
            sheet.range((row, col)).font.color = (0, 0, 0)
            sheet.range((row-2, col)).font.color = (255, 0, 0)
            sheet.range((row-1, col)).font.color = 0x3399FF # 藍色
            sheet.range((row+1, col)).font.color = 0x009900 # 綠色

            # 取得當前字元
            cell_value = strip_cell(v3_value[index])

            # 若為空白字元，則跳過
            if cell_value == None:
                index += 1
                msg = f"《空白字元》，跳過"
            elif cell_value == '\n':
                # 換行：列數加一，並從下一列的第一個字元開始
                sheet.range((row, col)).value = '=CHAR(10)'
                index += 1
                msg = f"《換行》"
            else:
                # 將【漢字】填入儲存格
                sheet.range((row, col)).value = cell_value
                index += 1
                msg = f"{cell_value}"

            print(f"{index}. ({row}, {col}) = {msg}")
            col += 1

            if cell_value == '\n':
                row += ROWS_PER_LINE
                col = start_col
            # 若已處理完整列(如：15字)，則換到下一列
            if col > end_col:
                col = start_col
                row += ROWS_PER_LINE
            # 若已處理完整個字串，則跳出迴圈
            if index == total_length:  break

        # 處理完一列後，換到下一列
        row += ROWS_PER_LINE
        col = start_col

    # 保存 Excel 檔案
    sheet.range((row, col)).value = "φ"
    wb.save()

    # 選擇名為 "顯示注音輸入" 的命名範圍
    named_range = wb.names['顯示注音輸入']
    named_range.refers_to_range.value = True

    logging.info(f"已將《文章》之漢字，填入【漢字注音】工作表之【漢字標音】工作區！")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 作業程序
# =========================================================================
def process(wb):
    logging_process_step("<----------- 作業開始！---------->")
    # ---------------------------------------------------------------------
    # 取得【待注音漢字】總字數
    # ---------------------------------------------------------------------
    cell_value = wb.sheets['漢字注音'].range('V3').value

    if cell_value is None:
        print("【待注音漢字】儲存格未填入文字，作業無法繼續。")
        logging.warning("【待注音漢字】儲存格為空")
        return EXIT_CODE_INVALID_INPUT

    value_length = len(cell_value.strip())
    print(f"【待注音漢字】總字數為: {value_length}")
    logging.info(f"【待注音漢字】總字數為: {value_length}")

    # ---------------------------------------------------------------------
    # 執行儲存格重設作業
    # ---------------------------------------------------------------------
    # print("清除儲存格內容...")
    # clear_han_ji_kap_piau_im(wb)
    # logging.info("儲存格內容清除完畢")

    # print("重設儲存格之格式...")
    # reset_cells_format_in_sheet(wb)
    # logging.info("儲存格格式重設完畢")

    status_code = fill_hanji_in_cells(wb)
    if status_code != EXIT_CODE_SUCCESS:
        logging_exc_error(msg="漢字填入標音作業儲存格失敗！", error=None)
        return EXIT_CODE_PROCESS_FAILURE
    else:
        logging_process_step("待注音漢字已填入【漢字注音】工作表")

    # ---------------------------------------------------------------------
    # 為漢字查找標音
    # ---------------------------------------------------------------------
    ue_im_lui_piat = wb.names['語音類型'].refers_to_range.value
    han_ji_khoo = wb.names['漢字庫'].refers_to_range.value

    if han_ji_khoo in ["河洛話", "廣韻"]:
        db_name = DB_HO_LOK_UE if han_ji_khoo == "河洛話" else DB_KONG_UN
        module_name = 'mod_河洛話' if han_ji_khoo == "河洛話" else 'mod_廣韻'

        try:
            # 查找漢字標音
            ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3',
                              ue_im_lui_piat=ue_im_lui_piat,
                              han_ji_khoo=han_ji_khoo, db_name=db_name,
                              module_name=module_name,
                              function_name='han_ji_ca_piau_im')
            logging_process_step("為【漢字】自動查找標音，作業已完成！")
        except Exception as e:
            logging_exc_error(msg="在查找漢字標音時發生錯誤！", error=e)
            return EXIT_CODE_PROCESS_FAILURE
    else:
        logging_exc_error(msg="無法作業，需請檢查【env】工作表之設定已確實完成！", error=None)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS

# =============================================================================
# 程式主流程
# =============================================================================
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
    # (1) 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
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
        logging_exc_error(f"無法找到作用中的 Excel 活頁簿！", error=e)
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
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=e)
            return EXIT_CODE_PROCESS_FAILURE

    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        #--------------------------------------------------------------------------
        # 儲存檔案
        #--------------------------------------------------------------------------
        try:
            # 要求畫面回到【漢字注音】工作表
            wb.sheets['漢字注音'].activate()
            # 儲存檔案
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging_exc_error(msg="儲存檔案失敗！", error=e)
                return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

        # if wb:
        #     xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

