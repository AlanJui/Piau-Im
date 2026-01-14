# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

from mod_ca_ji_tian import HanJiTian  # 新的查字典模組
from mod_excel_access import get_value_by_name
from mod_logging import init_logging, logging_exc_error, logging_process_step

# 載入自訂模組
from mod_程式 import ExcelCell, Program

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_SAVE_FAILURE = 3
EXIT_CODE_PROCESS_FAILURE = 10
EXIT_CODE_UNKNOWN_ERROR = 99

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

# =========================================================================
# 主要處理函數
# =========================================================================
def _get_active_cell_from_sheet(sheet, xls_cell: ExcelCell):
    """處理整個工作表"""
    program = xls_cell.program

    # 自【作用儲存格】取得【Excel 儲存格座標】(列,欄) 座標
    active_cell = sheet.api.Application.ActiveCell
    if active_cell:
        # 顯示【作用儲存格】位置
        active_row = active_cell.Row
        active_col = active_cell.Column
        active_col_name = xw.utils.col_name(active_col)
        print(f"作用儲存格：{active_col_name}{active_row}（{active_cell.Row}, {active_cell.Column}）")

        # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
        line_start_row = 3  # 第一行【標音儲存格】所在 Excel 列號: 3
        line_no = (active_row - line_start_row + 1) // program.ROWS_PER_LINE
        row = line_start_row + (line_no * program.ROWS_PER_LINE)
        col = active_cell.Column
        cell = sheet.range((row, col))
        cell.select()

        # 處理儲存格
        xls_cell.process_cell(cell, row, col)


def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        program = Program(wb=wb, args=args, hanji_piau_im_sheet='漢字注音')

        # 建立儲存格處理器
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 作業處理中
    #--------------------------------------------------------------------------
    try:
        # 處理工作表
        sheet_name = program.hanji_piau_im_sheet
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 取得【作用儲存格】
        _get_active_cell_from_sheet(sheet=sheet, xls_cell=xls_cell)

        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dicts()

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 處理作業結束
    #--------------------------------------------------------------------------
    print('\n')
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main(args):
    """主程式 - 從 Excel 呼叫或直接執行"""
    try:
        # 取得 Excel 活頁簿
        wb = None
        try:
            # 嘗試從 Excel 呼叫取得（RunPython）
            wb = xw.Book.caller()
        except Exception:
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
            exit_code = process(wb=wb, args=args)

            # 儲存檔案
            if exit_code == EXIT_CODE_SUCCESS:
                try:
                    wb.save()
                    file_path = wb.fullname
                    logging_process_step(f"儲存檔案至路徑：{file_path}")
                except Exception as e:
                    logging_exc_error(msg="儲存檔案異常！", error=e)
                    return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

            return exit_code

    except Exception:
        logging.exception("程式執行失敗")
        return EXIT_CODE_UNKNOWN_ERROR


def test_han_ji_tian():
    """測試 HanJiTian 類別"""
    print("=" * 70)
    print("測試 HanJiTian 查詢功能")
    print("=" * 70)

    try:
        # 初始化字典
        ji_tian = HanJiTian(DB_HO_LOK_UE)

        # 測試查詢
        test_chars = ["東", "西", "南", "北", "中"]

        for han_ji in test_chars:
            print(f"\n查詢漢字：{han_ji}")
            result = ji_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="白話音")

            if result:
                for item in result:
                    print(f"  台語音標：{item['台語音標']}, 常用度：{item.get('常用度', 'N/A')}, 說明：{item.get('摘要說明', 'N/A')}")
            else:
                print(f"  查無資料")

        print("\n" + "=" * 70)
        print("測試完成")
        print("=" * 70)

    except Exception as e:
        print(f"測試失敗：{e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description='缺字表修正後續作業程式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a000.py          # 執行一般模式
  python a000.py -new     # 建立新的字庫工作表
  python a000.py -test    # 執行測試模式
'''
        )
    parser.add_argument(
        '--test',
        action='store_true',
        help='執行測試模式',
    )
    parser.add_argument(
        '--new',
        action='store_true',
        help='建立新的標音字庫工作表',
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        test_han_ji_tian()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，返回代碼：{exit_code}")
            sys.exit(exit_code)