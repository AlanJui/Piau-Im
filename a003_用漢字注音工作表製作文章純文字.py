# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

from mod_excel_access import reset_cells_format_in_sheet

# 載入自訂模組
from mod_file_access import dump_txt_file

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
# 作業程序
# =========================================================================
def process(wb):
    """
    將 Excel 工作表中指定區域的漢字取出，儲存為一個純文字檔。
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

    # 將所有漢字寫入文字檔
    output_dir_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    output_file = 'tmp.txt'
    output_file_path = os.path.join(output_dir_path, output_file)
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(han_ji_text)
    logging_process_step(f"已成功將漢字輸出至檔案：{output_file}")

    # 螢幕 Dump 檔案內容
    dump_txt_file(output_file)

    # 回填【漢字注音】工作表之 V3 儲存格
    wb.sheets['漢字注音'].range('V3').value = han_ji_text
    logging_process_step("已將【漢字注音】工作表之內容，匯整成整篇【文章】之純文字，並回存 V3 儲存格！")

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
    # 確認 Excel 應用程式是否已啟動
    if xw.apps.count == 0:
        logging_process_step("程式異常終止：未檢測到 Excel 應用程式！")
        return EXIT_CODE_INVALID_INPUT
    # 確認是否有 Excel 活頁簿檔案已開啟
    if xw.apps.active.books.count == 0:
        logging_process_step("程式異常終止：未檢測到 Excel 活頁簿！")
        return EXIT_CODE_INVALID_INPUT

    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
        directory = Path(wb.fullname).parent
        logging_process_step(f"作用中活頁簿：{wb.name}")
        logging_process_step(f"目錄路徑：{directory}")
    except Exception as e:
        logging_exc_error(msg=f"程式異常終止：{program_name}", error=e)
        return EXIT_CODE_INVALID_INPUT

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_INVALID_INPUT

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
            file_name = '_working'
            output_dir_path = Path(wb.fullname).parent
            file_path = os.path.join(output_dir_path, f"{file_name}.xlsx")
            wb.save(file_path)
            print(f"以檔案名稱：【{file_name}.xlsx】，儲存於目錄路徑：{output_dir_path}！")
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
