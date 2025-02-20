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
                try:
                    wb.sheets[sheet_name].delete()
                except Exception as e:
                    logging_exc_error(msg=f"刪除工作表 {sheet_name} 失敗！", error=e)
                    continue  # 繼續刪除其他工作表
    except Exception as e:
        logging_exc_error(msg="刪除工作表失敗！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    msg = f'確認過不應存在要的工作表，的確沒有或已經刪除：{sheets_to_delete}'
    logging_process_step(msg)

    #----------------------------------------------------------------------
    # 將儲存格內的舊資料清除
    #----------------------------------------------------------------------
    sheet = wb.sheets['漢字注音']   # 選擇工作表
    sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

    total_rows = wb.names['每頁總列數'].refers_to_range.value if '每頁總列數' in [name.name for name in wb.names] else 120
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
    # 清空【env】工作表之設定： 根據 name_list 清除對應名稱所在儲存格的內容
    #----------------------------------------------------------------------
    sheet = wb.sheets['env']   # 選擇工作表
    sheet.activate()               # 將「env」工作表設為作用中工作表
    name_list = ['INPUT_FILE_PATH', 'FILE_NAME', 'TITLE', 'IMAGE_URL', '章節序號']
    for name in name_list:
        if name in [n.name for n in wb.names]:
            wb.names[name].refers_to_range.clear_contents()
    logging_process_step("【env】工作表之所有選項亦被清除！")


    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
def main(mode: str = "1"):
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
    if mode == "1":
        # 設定活頁簿檔案路徑
        file_path = os.path.join(project_root, "Template.xlsx")
        # 確認檔案是否存在
        if not os.path.exists(file_path):
            logging_process_step("程式異常終止：找不到檔案！")
            return EXIT_CODE_NO_FILE
        # 開啟活頁簿檔案
        try:
            wb = xw.Book(file_path)
            logging_process_step(f"已開啟活頁簿：{file_path}")
        except Exception as e:
            logging_exc_error(msg="程式異常終止：無法開啟活頁簿！", error=e)
            return EXIT_CODE_NO_FILE
    else:
        # 確認 Excel 應用程式是否已啟動
        if xw.apps.count == 0:
            logging_process_step("程式異常終止：未檢測到 Excel 應用程式！")
            return EXIT_CODE_INVALID_INPUT
        # 確認是否有 Excel 活頁簿檔案已開啟
        if xw.apps.active.books.count == 0:
            logging_process_step("程式異常終止：未檢測到 Excel 活頁簿！")
            return EXIT_CODE_INVALID_INPUT

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
            output_dir_path = str(Path(wb.fullname).parent)
            output_dir_path = "output7" if mode == "1" else output_dir_path
            file_path = os.path.join(project_root, output_dir_path, f"{file_name}.xlsx")
            wb.save(file_path)
            logging_process_step(f"以檔案名稱：【{file_name}.xlsx】，儲存於目錄路徑：{output_dir_path}！")
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
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "1"

    exit_code = main(mode)
    sys.exit(exit_code)
