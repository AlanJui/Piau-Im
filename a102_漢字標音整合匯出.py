# ========================================================================
# 程式名稱：a102_漢字標音整合匯出.py
# 程式說明：將 Excel 工作表中的漢字和標音整合匯出成格式化文字檔。
# 輸出格式：
#   《標題》
#   標音
#
#   漢字
#   標音
#
#   ...
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


def read_line_content(sheet, row, start_col, end_col):
    """
    讀取一列的內容（漢字或標音），直到遇到控制符號或行尾。

    Returns:
        tuple: (content, is_newline, is_eof)
        - content: 讀取到的文字內容
        - is_newline: 是否遇到換行符號
        - is_eof: 是否遇到結束符號
    """
    content = ""
    is_newline = False
    is_eof = False

    for col in range(start_col, end_col):
        cell_value = sheet.range((row, col)).value

        if cell_value == 'φ':       # 結束標示
            is_eof = True
            break
        elif cell_value == '\n':    # 換行標示
            is_newline = True
            break
        elif cell_value == None or cell_value == '':    # 空白
            continue
        else:                       # 正常內容
            content += cell_value + ' '

    return content.strip(), is_newline, is_eof


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb):
    """
    將 Excel 工作表中的漢字和標音整合輸出。
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

    # 設定起始及結束的【列】位址
    TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
    ROWS_PER_LINE = 4
    han_ji_start_row = 5  # 漢字從第5列開始
    piau_im_start_row = 6  # 標音從第6列開始
    end_row = han_ji_start_row + (TOTAL_LINES * ROWS_PER_LINE)
    line = 1

    # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
    CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
    start_col = 4  # D欄
    end_col = start_col + CHARS_PER_ROW  # R欄

    #--------------------------------------------------------------------------
    # 作業處理：讀取標題
    #--------------------------------------------------------------------------
    logging_process_step(f"開始【處理作業】...")

    # 讀取文章標題（從第5列第4欄開始，讀取整列組合標題）
    title_parts = []
    for col in range(start_col, end_col):
        cell_value = sheet.range((5, col)).value
        if cell_value == '\n' or cell_value == 'φ':
            # 遇到換行或結束符號，立即終止
            break
        if cell_value and cell_value != '':
            title_parts.append(str(cell_value))

    title = ''.join(title_parts)
    if title:
        output_text = f"{title}\n"
    else:
        output_text = ""

    print(f"讀取標題：{title}")

    EOF = False
    first_line = True  # 標記是否為第一行（標題行）

    #--------------------------------------------------------------------------
    # 逐列處理作業
    #--------------------------------------------------------------------------
    for idx in range(TOTAL_LINES):
        if EOF:
            print(f"\n========== 遇到結束符號，終止處理 ==========")
            break

        han_ji_row = han_ji_start_row + (idx * ROWS_PER_LINE)
        piau_im_row = piau_im_start_row + (idx * ROWS_PER_LINE)

        print(f"\n---------- 處理第 {line} 行 (漢字列={han_ji_row}, 標音列={piau_im_row}) ----------")

        # 讀取漢字
        han_ji_content, han_ji_newline, han_ji_eof = read_line_content(
            sheet, han_ji_row, start_col, end_col
        )

        # 讀取標音
        piau_im_content, piau_im_newline, piau_im_eof = read_line_content(
            sheet, piau_im_row, start_col, end_col
        )

        # 檢查是否結束
        if han_ji_eof or piau_im_eof:
            EOF = True

        # 如果有內容，則輸出
        if han_ji_content or piau_im_content:
            if first_line:
                # 第一行只輸出標音（標題的標音）
                if piau_im_content:
                    output_text += f"{piau_im_content}\n\n"
                first_line = False
            else:
                # 其他行輸出：漢字 + 換行 + 標音 + 空行
                if han_ji_content:
                    output_text += f"{han_ji_content}\n"
                if piau_im_content:
                    output_text += f"{piau_im_content}\n"
                output_text += "\n"

        # 處理換行
        if han_ji_newline or piau_im_newline:
            print(f"  換行標記")

        line += 1

    #--------------------------------------------------------------------------
    # 將結果寫入文字檔
    #--------------------------------------------------------------------------
    output_dir_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    output_file = 'formatted_output.txt'
    output_file_path = os.path.join(output_dir_path, output_file)

    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(output_text)

    logging_process_step(f"已成功將內容輸出至檔案：{output_file_path}")

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
