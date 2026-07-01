"""
a223_以作用儲存格之人工標音更新標音字庫.py V0.2.4

依據【】工作表，取得當前【作用儲存格】手動輸入之【人工標音】，調整自漢字典
資料庫查得之【台語音標】，並據此更新【漢字標音】儲存格內容。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import (
    excel_address_to_row_col,
    get_active_cell,
    get_active_cell_address,
    get_line_no_by_row,
    get_row_by_line_no,
)

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)
from mod_程式 import ExcelCell, Program

init_logging()

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
DB_KONG_UN = os.getenv("DB_KONG_UN", "Kong_Un.db")

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

# =============================================================================
# 作業主流程
# =============================================================================


def process(wb, args) -> int:
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        # --------------------------------------------------------------------------
        # 初始化 process config
        # --------------------------------------------------------------------------
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器
        # xls_cell = ExcelCell(program=program)
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        # --------------------------------------------------------------------------
        # 處理作業開始
        # --------------------------------------------------------------------------
        source_sheet_name = "漢字注音"
        jin_kang_piau_im_sheet_name = "人工標音字庫"
        piau_im_ji_khoo_sheet_name = "標音字庫"

        # ----------------------------------------------------------------------
        # 取得【作用儲存格】
        # ----------------------------------------------------------------------
        # 指定【漢字注音】工作表為【作用工作表】
        source_sheet = wb.sheets[source_sheet_name]
        source_sheet.activate()

        active_cell_address = get_active_cell_address()
        row, col = excel_address_to_row_col(active_cell_address)
        current_line_no = get_line_no_by_row(current_row_no=row)  # 計算行號
        jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = (
            get_row_by_line_no(current_line_no)
        )

        han_ji = source_sheet.range((han_ji_row, col)).value
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value
        han_ji_position = (han_ji_row, col)
        han_ji_cell = source_sheet.range((han_ji_row, col))

        print(
            f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}"
        )
        print(f"📌 漢字：{han_ji}")
        print(
            f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}"
        )

        # ----------------------------------------------------------------------
        # 自【漢字注音】工作表之【作用儲存格】取得【人工標音】
        # ----------------------------------------------------------------------
        tai_gi_im_piau, han_ji_piau_im = xls_cell.jin_kang_piau_im_cu_han_ji_piau_im(
            jin_kang_piau_im=jin_kang_piau_im
        )

        # 將【台語音標】和【漢字標音】寫入【漢字注音】工作表之【作用儲存格】
        han_ji_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        han_ji_cell.offset(+1, 0).value = han_ji_piau_im  # 漢字標音
        msg = f"{han_ji}： [{jin_kang_piau_im}] / [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
        print(f"✅ 已更新儲存格：{active_cell_address}，內容為：{msg}")

        # -------------------------------------------------------------------------
        # 將【漢字注音】工作表，存放在【作用儲存格】的【人工標音】，轉換成【漢字標音】後：
        # 1. 在【人工標音字庫】工作表，留下紀錄：漢字、台語音標（取自【作用儲存格】的新【台語音標】）、校正音標（="N/A"）、座標（取自【作用儲存格】的儲存格座標）
        # 2. 自【標音字庫】工作表，依對映到【漢字注音】工作表的【座標】，找到資料錄。然後更新【資料紀錄】的【台語音標】欄。
        # 3. 依找到的【資料紀錄】，取【座欄】欄內的【座標清單】，依各個【座標】值，更新【漢字注音】工作表的【台語音標】及【漢字標音】。
        # -------------------------------------------------------------------------
        # 調整 row 指向【漢字】儲存格所在座標列
        row = han_ji_row

        # -------------------------------------------------------------------------
        # 在【人工標音字庫】工作表對映之【字庫】(dict)，添加或更新一筆【漢字】及
        # 【台語音標】資料：依【漢字】及【座標】在【人工標音字庫】工作表查檢是否已
        # 有紀錄，若有：執行【更新】作業；若無：執行【添加】作業。
        # -------------------------------------------------------------------------
        xls_cell.jin_kang_piau_im_ji_khoo_dict.add_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau="N/A",
            coordinate=(row, col),
        )
        # 將【人工標音字庫】，寫入 Excel 工作表
        xls_cell.jin_kang_piau_im_ji_khoo_dict.write_to_excel_sheet(
            wb=wb, sheet_name=jin_kang_piau_im_sheet_name
        )
        # -------------------------------------------------------------------------
        # 取得該【座標】目前所屬之【標音字庫】資料紀錄（本例中為【台語音標】="ti1" 的
        # 紀錄），並取出其完整之【座標】清單。這份清單代表：在人工訂正之前，此漢字在
        # 【漢字注音】工作表中，凡是與【作用儲存格】共用同一（待訂正）讀音的所有出現
        # 位置——這些位置理應隨同校正為這次輸入的新【台語音標】。
        # -------------------------------------------------------------------------
        old_entry = xls_cell.piau_im_ji_khoo_dict.get_entry_by_han_ji_and_coordinate(
            han_ji=han_ji, coordinate=(row, col)
        )
        old_coord_list = (
            list(old_entry.get("coordinates", [])) if old_entry else [(row, col)]
        )

        # -------------------------------------------------------------------------
        # 【注意】：【標音字庫】的【座標】欄，須完整保留所有指向【漢字注音】工作表的
        # 紀錄，絕對不可移除既有座標。因此本處【不】自舊【台語音標】資料紀錄中移除
        # 任何座標，僅將上述取得的【座標清單】，逐一併入登錄至新【台語音標】的資料
        # 紀錄（若尚無此【台語音標】之紀錄，add_or_update_entry() 會自動新增一筆）。
        # -------------------------------------------------------------------------
        for coord in old_coord_list:
            xls_cell.piau_im_ji_khoo_dict.add_or_update_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                hau_ziann_im_piau="N/A",
                coordinates=coord,
            )
        # 將【標音字庫】，寫入 Excel 工作表
        xls_cell.piau_im_ji_khoo_dict.write_to_excel_sheet(
            wb=wb, sheet_name=piau_im_ji_khoo_sheet_name
        )

        # -------------------------------------------------------------------------
        # 依據【座標】清單，指向【漢字注音】工作表各【漢字】儲存格的【台語音標】及【漢字標音】
        # 《注音》：使用【標音字庫】工作表的【座標】資料欄中的【座標清單】時，應剔除【人工標音字庫】
        # 已標注的【座標】。
        # -------------------------------------------------------------------------
        # 取得【標音字庫】中，此【漢字】+【台語音標】對映的【座標清單】
        piau_im_coord_list = (
            xls_cell.piau_im_ji_khoo_dict.get_coordinates_by_han_ji_and_tai_gi_im_piau(
                han_ji=han_ji, tai_gi_im_piau=tai_gi_im_piau
            )
        )

        # 取得【人工標音字庫】中，此【漢字】已登錄之全部【座標】（跨所有音讀項目），
        # 用以自【座標清單】中剔除，避免重複覆寫已由人工標音掌管之儲存格
        jin_kang_coord_list = []
        for jin_kang_entry in xls_cell.jin_kang_piau_im_ji_khoo_dict.ji_khoo_dict.get(
            han_ji, []
        ):
            jin_kang_coord_list.extend(jin_kang_entry.get("coordinates", []))

        # 剔除【人工標音字庫】已標注的座標後，將【標音字庫】座標清單中剩餘的座標，
        # 對映到【漢字注音】工作表的【漢字】儲存格，同步更新其上、下方之
        # 【台語音標】、【漢字標音】儲存格內容
        for coord_row, coord_col in piau_im_coord_list:
            if (coord_row, coord_col) in jin_kang_coord_list:
                continue
            target_han_ji_cell = source_sheet.range((coord_row, coord_col))
            target_han_ji_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
            target_han_ji_cell.offset(+1, 0).value = han_ji_piau_im  # 漢字標音
            print(
                f"✅ 同步更新位於【漢字注音】工作表的儲存格：({coord_row}, {coord_col}) ==> "
                f"台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}"
            )

        # -------------------------------------------------------------------------
        # 更新資料庫中【漢字庫】資料表
        # -------------------------------------------------------------------------
        siong_iong_too_to_use = (
            0.8 if program.ue_im_lui_piat == "文讀音" else 0.6
        )  # 根據語音類型設定常用度
        xls_cell.insert_or_update_to_db(
            table_name=program.table_name,
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            ue_im_lui_piat=program.ue_im_lui_piat,
            siong_iong_too=siong_iong_too_to_use,
        )

        logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
        # 再次拋出異常，讓外層函式能捕捉
        raise


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
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
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        exit_code = process(wb, args)
    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 要求畫面回到【漢字注音】工作表
        # wb.sheets['漢字注音'].activate()
        # 儲存檔案
        wb.save()
        file_path = wb.fullname
        logging_process_step(f"儲存檔案至路徑：{file_path}")

    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束作業
    # =========================================================================
    return EXIT_CODE_SUCCESS


def ut01():
    # 取得【作用中活頁簿】
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="缺字表修正後續作業程式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python a000.py          # 執行一般模式
  python a000.py -new     # 建立新的字庫工作表
  python a000.py -test    # 執行測試模式
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        ut01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code == EXIT_CODE_SUCCESS:
            print("程式正常完成！")
        else:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
