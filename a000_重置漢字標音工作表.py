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
# 作業程序
# =========================================================================
def process(wb):
    logging_process_step("<----------- 作業開始！---------->")
    #----------------------------------------------------------------------
    # 刪除漢字標音作業中使用之工作表
    #----------------------------------------------------------------------
    # 要刪除的工作表名稱
    sheets_to_delete = ["缺字表", "標音字庫", "人工標音字庫"]

    try:
        for sheet_name in sheets_to_delete:
            # 如果工作表確實存在才刪除
            if sheet_name in [sh.name for sh in wb.sheets]:
                wb.sheets[sheet_name].delete()
    except Exception as e:
        logging_exc_error(msg="刪除工作表失敗！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    msg = f'已刪除不必要的工作表：{sheets_to_delete}'
    logging_process_step(msg)

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
    reset_cells_format_in_sheet(wb)
    logging_process_step("【漢字注音】工作表，完成重置作業！")

    #--------------------------------------------------------------------------
    # 清空【env】工作表之設定
    #--------------------------------------------------------------------------
    sheet = wb.sheets['env']   # 選擇工作表
    sheet.activate()               # 將「env」工作表設為作用中工作表
    end_of_row = 20
    sheet.range(f'C2:C{end_of_row}').clear_contents()     # 清除 C3:R{end_of_row} 範圍的內容
    logging_process_step("【env】工作表之所有選項亦被清除！")

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
