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
from mod_excel_access import save_as_new_file  # noqa: F401

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
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


# =============================================================================
# 作業主流程
# =============================================================================
def _show_separtor_line(source_sheet_name: str, target_sheet_name: str):
    print('\n\n')
    print("=" * 100)
    print(f"使用【{source_sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
    print("=" * 100)


def _process_sheet(sheet, program: Program, xls_cell: ExcelCell):
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(program.config.start_col)}{program.config.line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, program.TOTAL_LINES + 1):
        if is_eof: break  # noqa: E701
        line_no = r
        print('-' * 60)
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
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = xls_cell._process_cell(active_cell, row, col)
            if new_line: break  # noqa: E701
            if is_eof: break  # noqa: E701


def process(wb, args) -> int:
    """
    將 Excel 工作表中的漢字和標音整合輸出。
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        program = Program(wb, args, hanji_piau_im_sheet='漢字注音')

        # 建立儲存格處理器
        # xls_cell = ExcelCell(program=program)
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        #--------------------------------------------------------------------------
        # 處理作業開始
        #--------------------------------------------------------------------------
        # 處理工作表
        sheet_name = '漢字注音'
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        piau_im_ji_khoo_dict = xls_cell.piau_im_ji_khoo_dict
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
            print(f"開始測試【{sheet.name}】：作用儲存格：{active_cell.address}，漢字：{han_ji}，台語音標：{tai_gi_im_piau}")

            # print(f"作用儲存格：{active_cell.address}，漢字：{han_ji}")
            # tai_gi_im_piau = piau_im_ji_khoo_dict.get_tai_gi_im_piau_by_han_ji(han_ji=han_ji)
            # print(f"標音字庫查到的台語音標：{tai_gi_im_piau}")

            row_no = piau_im_ji_khoo_dict.get_row_by_han_ji_and_tai_gi_im_piau(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau
            )
            print(f"{han_ji}（{tai_gi_im_piau}）落在【標音字庫】的 Row 號：{row_no}")

            # 依【漢字】與【台語音標】取得在【標音字庫】工作表中的【座標】清單
            coord_list = piau_im_ji_khoo_dict.get_coordinates_by_han_ji_and_tai_gi_im_piau(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau
            )
            print(f"在【標音字庫】工作表，{han_ji}（{tai_gi_im_piau}）的座標清單：{coord_list}")

            # 檢驗(row, col)座標，是否在座標清單中
            coord_to_remove = (row, col)
            if coord_to_remove in coord_list:
                print(f"座標 {coord_to_remove} 有在座標清單之中。")
                # 刪除座標作業
                piau_im_ji_khoo_dict.remove_coordinate(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    coordinate=coord_to_remove
                )
                print(f"已從【標音字庫】工作表，移除 {han_ji}（{tai_gi_im_piau}）的座標：{coord_to_remove} ...")
            else:
                print(f"座標 {coord_to_remove} 不在座標清單之中。")

            # 儲存回 Excel
            print("將更新後的【標音字庫】寫回 Excel 工作表...")
            piau_im_ji_khoo_dict.write_to_excel_sheet(
                wb=wb,
                sheet_name='標音字庫'
            )

        def _test_update_entry_in_ji_khoo_dict():
            print("開始測試 update_entry_in_ji_khoo_dict() 方法...")
            row = 5
            col = 6
            # 設定作用儲存格
            source_sheet_name = '漢字注音'
            source_sheet = wb.sheets[source_sheet_name]
            active_cell = source_sheet.range((row, col))
            active_cell.select()
            han_ji = active_cell.value
            tai_gi_im_piau = active_cell.offset(-1, 0).value

            target_sheet_name = '標音字庫'
            msg = f"（{row}, {col}）：{han_ji}【{tai_gi_im_piau}】"
            print(f"更新【{target_sheet_name}】，作用儲存格：{msg}")

            xls_cell.update_hanji_zu_im_sheet_by_khuat_ji_piau(
                source_sheet_name=source_sheet_name,
                target_sheet_name=target_sheet_name,
            )

        def _test_copy_entry_from_jin_kang_piau_im_ji_khoo_dict():
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
                is_eof, new_line = xls_cell._process_cell(active_cell, row, col)

            # 處理工作表整列
            print("開始測試 _test_copy_entry_from_jin_kang_piau_im_ji_khoo_dict() 方法...")
            sheet = wb.sheets['漢字注音']
            _test_process_sheet(
                sheet=sheet,
                program=program,
                xls_cell=xls_cell,
            )

        def _test_new_entry_into_jin_kang_piau_im_ji_khoo_dict(row: int = 5, col: int = 6):
            def _test_process_sheet(sheet, program: Program, xls_cell: ExcelCell):
                # 設定測試環境
                row = 5
                col = 6
                active_cell = sheet.range((row, col))
                # 處理儲存格
                han_ji = active_cell.value
                tai_gi_im_piau = active_cell.offset(-1, 0).value
                target = f"（{row}, {col}）= {xw.utils.col_name(col)}{row} ==> {han_ji}【{tai_gi_im_piau}】"
                print("開始測試 process_cell() 方法...")
                print(f"作用儲存格：{target}")
                # 模擬處理作用儲存格
                is_eof = False
                new_line = False
                # 處理儲存格
                print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
                is_eof, new_line = xls_cell._process_cell(active_cell, row, col)

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

    except Exception:
        logging.exception("自動為【漢字】查找【台語音標】作業，發生例外！")
        raise

        # 逐列處理
        _process_sheet(
            sheet=sheet,
            program=program,
            xls_cell=xls_cell,
        )
    except Exception as e:
        msg=f"處理作業，發生異常！ ==> error = {e}"
        logging.exception(msg)
        raise

    #--------------------------------------------------------------------------
    # 處理作業結束
    #--------------------------------------------------------------------------
    # 寫回字庫到 Excel
    xls_cell.save_all_piau_im_ji_khoo_dict()

    print('\n')
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


def main():
    """主程式"""
    try:
        # 取得 Excel 活頁簿
        wb = None
        # 若失敗，則取得作用中的活頁簿
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
            return EXIT_CODE_NO_FILE

        if not wb:
            logging.error("無法取得 Excel 活頁簿")
            return EXIT_CODE_NO_FILE

        # 執行處理
        exit_code = process(wb=wb, args=None)

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
    sys.exit(exit_code)