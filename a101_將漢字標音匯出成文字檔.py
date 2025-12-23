# ========================================================================
# 程式名稱：a101_將漢字標音匯出成文字檔.py
# 程式說明：將 Excel 工作表中【漢字標音】儲存格的標音取出，儲存為一個純文字檔。
# 1. 讀取儲存格位置：從第 6 列開始（漢字標音列），每隔 4 列讀取一次（6, 10, 14, ...）
# 2. 欄位範圍：D欄到R欄（根據「每列總字數」設定）
# 3. 輸出檔案：儲存為 piau_im.txt
# 4. 特殊字元處理：
#    - φ：文字終結標示
#    - \n：換行標示
#    - None：空白儲存格（連續兩個空白視為結尾）
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
    format='%(asctime)s - %(levelname)s - %(levelname)s - %(message)s'
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
# Local Function
# =========================================================================
def dump_txt_file(file_path):
    """
    在螢幕 Dump 純文字檔內容。
    """
    print("\n【文字檔內容】：")
    print("========================================\n")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            print(content)
    except FileNotFoundError:
        print(f"無法找到檔案：{file_path}")

# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb):
    """
    將 Excel 工作表中【漢字標音】儲存格的標音取出，儲存為一個純文字檔。
    讀取儲存格：6D:6R, 10D:10R, 14D:14R, ...
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業開始！---------->")

    # 選擇工作表
    sheet = wb.sheets['漢字注音']
    sheet.activate()

    #--------------------------------------------------------------------------
    # 自【env】設定工作表，取得處理作業所需參數
    #--------------------------------------------------------------------------

    # 設定起始及結束的【列】位址（【第6列】、【第10列】、【第14列】等列）
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    start_row = 6  # 【漢字標音】列從第6列開始
    end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)
    line = 1

    # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start_col = 4  # D欄
    end_col = start_col + CHARS_PER_ROW  # R欄

    #--------------------------------------------------------------------------
    # 作業處理：逐列取出漢字標音，組合成純文字檔
    #--------------------------------------------------------------------------
    logging_process_step(f"開始【處理作業】...")
    piau_im_text = ""
    EOF = False

    # 逐列處理作業
    for row in range(start_row, end_row, ROWS_PER_LINE):
        # 若已到【結尾】或【超過總行數】，則跳出迴圈
        if EOF:
            print(f"\n========== 遇到結束符號，終止處理 ==========")
            break
        if line > TOTAL_LINES:
            print(f"\n========== 已處理 {line-1} 行，達到總行數 {TOTAL_LINES}，終止處理 ==========")
            break

        print(f"\n---------- 處理第 {line} 行 (row={row}) ----------")

        # 設定【作用儲存格】為列首
        sheet.range((row, 1)).select()

        # 計算對應的漢字列（標音列的前一列）
        han_ji_row = row - 1

        # 逐欄取出標音處理
        for col in range(start_col, end_col):
            # 先檢查對應的漢字儲存格是否有控制符號
            han_ji_value = sheet.range((han_ji_row, col)).value

            if han_ji_value == 'φ':       # 讀到【結尾標示】
                EOF = True
                col_name = xw.utils.col_name(col)
                print(f"({han_ji_row}, {col_name}) [漢字列] = 【文字終結】")
                break  # 立即跳出內層迴圈
            elif han_ji_value == '\n':    # 讀到【換行標示】
                piau_im_text += '\n'
                col_name = xw.utils.col_name(col)
                print(f"({han_ji_row}, {col_name}) [漢字列] = 【換行】")
                break  # 立即跳出內層迴圈

            # 取得當前標音儲存格內含值
            cell_value = sheet.range((row, col)).value

            if cell_value == None or cell_value == '':    # 讀到【空白】
                # 標音列可能有空白儲存格（漢字沒有標音），不加入文字但繼續處理
                continue  # 跳過此儲存格，不顯示也不中斷處理
            else:                       # 讀到：標音
                piau_im_text += cell_value + ' '  # 每個音節後面加空白
                # 顯示處理進度
                col_name = xw.utils.col_name(col)
                print(f"({row}, {col_name}) = {cell_value}")

        # 每當處理一行 15 個標音後，亦換到下一行
        line += 1

    # 將所有標音寫入文字檔
    output_dir_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    output_file = 'piau_im.txt'
    output_file_path = os.path.join(output_dir_path, output_file)
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(piau_im_text)
    logging_process_step(f"已成功將標音輸出至檔案：{output_file_path}")

    # 螢幕 Dump 檔案內容
    dump_txt_file(output_file_path)

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
