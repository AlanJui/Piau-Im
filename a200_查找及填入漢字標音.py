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
from mod_ca_ji_tian import HanJiTian
from mod_excel_access import save_as_new_file

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
# 作業處理函數
# =========================================================================
def _show_separtor_line(source_sheet_name: str, target_sheet_name: str):
    print('\n\n')
    print("=" * 100)
    print(f"使用【{source_sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
    print("=" * 100)

# =============================================================================
# 作業主流程
# =============================================================================
def process_sheet(sheet, program: Program, xls_cell: ExcelCell):
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(program.start_col)}{program.line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, program.TOTAL_LINES + 1):
        if is_eof: break
        line_no = r
        print('=' * 80)
        print(f"處理第 {line_no} 行...")
        row = program.line_start_row + (r - 1) * program.ROWS_PER_LINE + program.han_ji_row_offset
        new_line = False
        for c in range(program.start_col, program.end_col + 1):
            if is_eof: break
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()
            # 處理儲存格
            print('-' * 80)
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            is_eof, new_line = xls_cell.process_cell(active_cell, row, col)
            if new_line: break  # noqa: E701
            if is_eof: break  # noqa: E701


def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件

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

        # 逐列處理
        process_sheet(
            sheet=sheet,
            program=program,
            xls_cell=xls_cell,
        )

        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dict()

        #--------------------------------------------------------------------------
        # 處理作業結束
        #--------------------------------------------------------------------------
        print('=' * 80)
        logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(msg=f"在【{sheet_name}】工作表，自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
        raise

# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    """主程式 - 從 Excel 呼叫或直接執行"""
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
    # 取得【作用中活頁簿】
    wb = None
    try:
        # 嘗試從 Excel 呼叫取得（RunPython）
        wb = xw.Book.caller()
    except Exception:
        # 若失敗，則取得作用中的活頁簿
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging_exc_error(msg=f"無法找到作用中的 Excel 工作簿！", error=e)
            return EXIT_CODE_NO_FILE

    if not wb:
        logging_exc_error(msg="無法取得 Excel 活頁簿！")
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    exit_code = process(wb, args)

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    if exit_code == EXIT_CODE_SUCCESS:
        file_path = save_as_new_file(wb=wb)
        if not file_path:
            logging_exc_error(msg="儲存檔案失敗！", error=None)
            exit_code = EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
        else:
            logging_process_step(f"儲存檔案至路徑：{file_path}")

    # =========================================================================
    # 結束程式
    # =========================================================================
    print('\n')
    print('=' * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    if exit_code == EXIT_CODE_SUCCESS:
        return EXIT_CODE_SUCCESS    # 作業正常結束
    else:
        msg = f"程式異常終止，返回失敗碼：{exit_code}"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE


def test_01():
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
        description='依【漢字】查找【台語音標】並轉換成【漢字標音】',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a200_查找及填入漢字標音.py          # 執行一般模式
  python a200_查找及填入漢字標音.py -new     # 建立新的字庫工作表
  python a200_查找及填入漢字標音.py -test    # 執行測試模式
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
        help='建立新的字庫工作表',
    )
    args = parser.parse_args()
    new_piau_im_sheets = args.new
    test_mode = args.test

    if test_mode:
        # 執行測試
        test_01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        # 只在命令列執行時使用 sys.exit()，避免在調試環境中引發 SystemExit 例外
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)