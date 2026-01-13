# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sys

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_程式 import ExcelCell, Program

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

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
# 設定日誌
# =========================================================================
from mod_logging import (  # noqa: E402
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)

init_logging()


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def _process_sheet(sheet, program: Program, xls_cell: ExcelCell):
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(program.start_col)}{program.line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, program.TOTAL_LINES + 1):
        if is_eof: break  # noqa: E701
        line_no = r
        print('=' * 80)
        print(f"處理第 {line_no} 行...")
        row = program.line_start_row + (r - 1) * program.ROWS_PER_LINE + program.han_ji_row_offset
        new_line = False
        for c in range(program.start_col, program.end_col + 1):
            if is_eof: break  # noqa: E701
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()
            # 處理儲存格
            print('-' * 60)
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = xls_cell.process_cell(active_cell, row, col)
            if new_line: break  # noqa: E701
            if is_eof: break  # noqa: E701


def process(wb, args) -> int:
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業開始
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    try:
        program = Program(wb, args, hanji_piau_im_sheet='漢字注音')

        # 建立儲存格處理器
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        # --------------------------------------------------------------------------
        # 測試
        # --------------------------------------------------------------------------
        def _test100(row: int = 5, col: int = 6):
            # 設定作用儲存格
            sheet = wb.sheets['漢字注音'].activate()
            # active_cell = sheet.range('F5')
            # active_cell = wb.sheets['漢字注音'].range('F5')
            active_cell = wb.sheets['漢字注音'].range((row, col))
            active_cell.select()
            han_ji = active_cell.value
            tai_gi_im_piau = active_cell.offset(-1, 0).value
            print(f"開始測試【{sheet.name}】工作表：作用儲存格：{active_cell.address}，漢字：{han_ji}，台語音標：{tai_gi_im_piau}")

            # print(f"作用儲存格：{active_cell.address}，漢字：{han_ji}")
            # tai_gi_im_piau = piau_im_ji_khoo_dict.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            # print(f"標音字庫查到的台語音標：{tai_gi_im_piau}")

            row_no = xls_cell.piau_im_ji_khoo_dict.get_row_by_han_ji_and_tai_gi_im_piau(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau
            )
            print(f"{han_ji}（{tai_gi_im_piau}）落在【標音字庫】的 Row 號：{row_no}")

            # 依【漢字】與【台語音標】取得在【標音字庫】工作表中的【座標】清單
            coord_list = xls_cell.piau_im_ji_khoo_dict.get_coordinates_by_han_ji_and_tai_gi_im_piau(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau
            )
            print(f"在【標音字庫】工作表，{han_ji}（{tai_gi_im_piau}）的座標清單：{coord_list}")

            # 檢驗(row, col)座標，是否在座標清單中
            coord_to_remove = (row, col)
            if coord_to_remove in coord_list:
                print(f"座標 {coord_to_remove} 有在座標清單之中。")
                # 刪除座標作業
                xls_cell.piau_im_ji_khoo_dict.remove_coordinate(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    coordinate=coord_to_remove
                )
                print(f"已從【標音字庫】工作表，移除 {han_ji}（{tai_gi_im_piau}）的座標：{coord_to_remove} ...")
            else:
                print(f"座標 {coord_to_remove} 不在座標清單之中。")

            # 儲存回 Excel
            print("將更新後的【標音字庫】寫回 Excel 工作表...")
            xls_cell.piau_im_ji_khoo_dict.write_to_excel_sheet(
                wb=wb,
                sheet_name='標音字庫'
            )

        def _test_update_entry_in_ji_khoo_dict():
            print("開始測試 update_entry_in_ji_khoo_dict() 方法...")
            row = 5
            col = 6
            # 設定作用儲存格
            sheet = wb.sheets['漢字注音'].activate()
            active_cell = wb.sheets['漢字注音'].range((row, col))
            active_cell.select()
            han_ji = active_cell.value
            tai_gi_im_piau = active_cell.offset(-1, 0).value

            target = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】"
            ji_khoo_name = '標音字庫'
            print(f"更新【{sheet.name}】工作表對映之【{ji_khoo_name}】，作用儲存格：{target}")

            # TODO: Replace with the correct method name from ExcelCell class
            # xls_cell.update_entry_into_piau_im_dict(
            #     wb=wb,
            #     target_dict=xls_cell.piau_im_ji_khoo_dict,
            #     han_ji=han_ji,
            #     tai_gi_im_piau=tai_gi_im_piau,
            #     hau_ziann_im_piau='N/A',
            #     row=row,  col=col
            # )
            print("Method not implemented - check ExcelCell class for correct method name")

        def _test_process_sheet(sheet, program: Program, xls_cell: ExcelCell):
            # 設定測試環境
            row = 17
            col = 7
            active_cell = sheet.range((row, col))
            # 處理儲存格
            han_ji = active_cell.value
            tai_gi_im_piau = active_cell.offset(-1, 0).value
            jin_kang_piau_im = active_cell.offset(-2, 0).value
            target = f"（{row}, {col}）= {xw.utils.col_name(col)}{row} ==> {han_ji}【{tai_gi_im_piau}】({jin_kang_piau_im})"
            print("開始測試 process_cell() 方法...")
            print(f"作用儲存格：{target}")
            # 模擬處理作用儲存格
            is_eof = False
            new_line = False
            # 處理儲存格
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = xls_cell.process_cell(active_cell, row, col)

        def _test_copy_entry_from_jin_kang_piau_im_ji_khoo_dict():
            # 處理工作表整列
            print("開始測試 _test_copy_entry_from_jin_kang_piau_im_ji_khoo_dict() 方法...")
            sheet = wb.sheets['漢字注音']
            _test_process_sheet(
                sheet=sheet,
                program=program,
                xls_cell=xls_cell,
            )

        def _test_new_entry_into_jin_kang_piau_im_ji_khoo_dict(row: int = 5, col: int = 6):
            # 處理工作表整列
            print("開始測試 _test_new_entry_into_jin_kang_piau_im_ji_khoo_dict() 方法...")
            sheet = wb.sheets['漢字注音']
            _test_process_sheet(
                sheet=sheet,
                program=program,
                xls_cell=xls_cell,
            )

        def _test_normaal_mode():
            """測試一般模式"""
            sheet_name = '漢字注音'
            sheet = wb.sheets[sheet_name]
            sheet.activate()

            # 逐列處理
            _process_sheet(
                sheet=sheet,
                program=program,
                xls_cell=xls_cell,
            )
        # --------------------------------------------------------------------------
        # 測試作業
        # --------------------------------------------------------------------------
        # _test_normaal_mode()
        # _test100()
        # _test_update_entry_in_ji_khoo_dict()
        # _test_new_entry_into_jin_kang_piau_im_ji_khoo_dict()
        # _test_copy_entry_from_jin_kang_piau_im_ji_khoo_dict()
        _test_new_entry_into_jin_kang_piau_im_ji_khoo_dict()

        print('=' * 40)
        print("測試結束。")
        print('=' * 40)

        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dict()

        return EXIT_CODE_SUCCESS

    except Exception:
        logging.exception("自動為【漢字】查找【台語音標】作業，發生例外！")
        raise


def main():
    """主程式"""
    try:
        # 取得 excel 活頁簿
        wb = None
        # 若失敗，則取得作用中的活頁簿
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging_exc_error(msg="無法找到作用中的 excel 工作簿！", error=e)
            return EXIT_CODE_NO_FILE

        if not wb:
            logging_exc_error(msg="無法取得 excel 活頁簿！",
                              error=Exception("No active workbook found"))
            return EXIT_CODE_NO_FILE

        # 執行處理
        exit_code = process(wb, sys.argv)
        return exit_code

    except FileNotFoundError as fnf_error:
        logging_exception(msg="找不到指定的檔案！", error=fnf_error)
        return EXIT_CODE_NO_FILE
    except ValueError as val_error:
        logging_exception(msg="輸入資料有誤！", error=val_error)
        return EXIT_CODE_INVALID_INPUT
    except Exception as e:
        logging_exception(msg="處理過程中發生未知錯誤！", error=e)
        return EXIT_CODE_UNKNOWN_ERROR


if __name__ == "__main__":
    import sys

    exit_code = main()
    if exit_code != EXIT_CODE_SUCCESS:
        sys.exit(exit_code)
