# a200_查找及填入漢字標音.py v0.2.2.3
# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# 載入自訂模組/函式
from mod_ca_ji_tian import HanJiTian
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)
from mod_標音 import is_han_ji
from mod_程式 import ExcelCell, Program

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
init_logging()

# =========================================================================
# 作業協助函數
# =========================================================================
def _show_separtor_line(source_sheet_name: str, target_sheet_name: str):
    print('\n\n')
    print("=" * 100)
    print(f"使用【{source_sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
    print("=" * 100)

# =========================================================================
# 資料類別：儲存處理配置
# =========================================================================
class CellProcessor(ExcelCell):
    """
    個人字典查詢專用的儲存格處理器
    繼承自 ExcelCell
    覆蓋以下方法以實現個人字典查詢功能：
    - _process_sheet(): 處理整個工作表
    """
    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        """
        初始化處理器
        :param config: 設定檔物件 (包含標音方法、資料庫連線等)
        :param jin_kang_ji_khoo: 人工標音字庫 (JiKhooDict) - 用於 '=' 查找
        :param piau_im_ji_khoo: 標音字庫
        :param khuat_ji_piau_ji_khoo: 缺字表
        """
        # 調用父類別（MengDianExcelCell）的建構子
        super().__init__(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

    #------------------------------------------------------------------
    # 覆蓋父類別之【方法】（method）
    #------------------------------------------------------------------
    def _process_cell(
        self,
        cell,
        row: int,
        col: int,
    ) -> int:
        """
        處理單一儲存格

        Returns:
            status_code: 儲存格內容代碼
                0 = 漢字
                1 = 文字終結符號
                2 = 換行符號
                3 = 空白、標點符號等非漢字字元
        """
        # 初始化樣式
        self._reset_cell_style(cell)

        # 取得【漢字】儲存格內容
        cell_value = cell.value
        print("-" * 40)

        # 檢查是否有【人工標音】
        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            self._process_jin_kang_piau_im(
                han_ji=cell_value,
                jin_kang_piau_im=jin_kang_piau_im,
                cell=cell,
                row=row,
                col=col,
            )
            return 0  # 漢字

        # 依據【漢字】儲存格內容進行處理
        if cell_value == 'φ':
            self._show_msg(row, col, "【文字終結】")
            return  1   # 文章終結符號
        elif cell_value == '\n':
            self._show_msg(row, col, "【換行】")
            return  2   #【換行】
        elif not is_han_ji(cell_value):
            if cell_value is None or str(cell_value).strip() == "":
                self._show_msg(row, col, "【空白】")
            else:
                self._show_msg(row, col, cell_value)
                self._process_non_han_ji(cell_value)
            return 3    # 空白或標點符號
        else:
            self._show_msg(row, col, cell_value)
            self._process_han_ji(cell_value, cell, row, col)
            return  0  # 漢字

    def _process_sheet(self, sheet):
        """處理整個工作表"""
        # 初始化變數
        config = self.program
        total_lines = config.TOTAL_LINES
        rows_per_line = config.ROWS_PER_LINE
        line_start_row = config.line_start_row
        # start_row = line_start_row + 2  # 調整為實際起始列
        # end_row = start_row + (config.TOTAL_LINES * config.ROWS_PER_LINE)
        start_col = config.start_col
        end_col = config.end_col
        han_ji_row_offset = config.han_ji_row_offset

        #--------------------------------------------------------------------------
        # 處理作用中列(row)的所有儲存格
        #--------------------------------------------------------------------------
        active_cell = sheet.range(f'{xw.utils.col_name(start_col)}{line_start_row}')
        active_cell.select()

        is_eof = False
        for line_no in range(1, total_lines + 1):
            # 檢查是否到達結尾
            if is_eof or line_no > total_lines:
                break

            # 初始化每行所需使用變數
            is_eol = False

            # 顯示目前處理【第 n 行】
            self._show_separtor_line(f"處理第 {line_no} 行...")

            # 調整 row 值至【漢字】儲存格所在列
            # （每【行（line）】由 4【列（row）】所構成，漢字在第 3 列：5, 9, 13, ... ）
            row = line_start_row + (line_no - 1) * rows_per_line + han_ji_row_offset

            #----------------------------------------------------------------------
            # 處理列中所有欄(col)儲存格
            #----------------------------------------------------------------------
            for c in range(start_col, end_col + 1):
                # 初始化每列所需使用變數
                status_code = 0

                # 將目前處理之儲存格，設為作用中儲存格
                row = row
                col = c
                active_cell = sheet.range((row, col))
                active_cell.select()

                # 顯示正要處理的儲存格座標位置
                print('-' * 80)
                print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")

                #------------------------------------------------------------------
                # 處理儲存格
                #------------------------------------------------------------------
                # status_code:
                # 0 = 儲存格內容為：漢字
                # 1 = 儲存格內容為：文字終結符號
                # 2 = 儲存格內容為：換行符號
                # 3 = 儲存格內容為：空白、標點符號等非漢字字元
                status_code = self._process_cell(active_cell, row, col)

                # 檢查是否需因：換行、文章終結，而跳出內層迴圈
                if status_code == 1:
                    is_eof = True
                    break
                elif status_code == 2:
                    is_eol = True
                    break

                # 檢查處理作業【是否已達行尾】或【讀到換行符號】
                if is_eol or col == end_col - 1:
                    print('\n')
                    print('=' * 60)
                    print('\n')

        # 將字庫 dict 回存 Excel 工作表
        self.save_all_piau_im_ji_khoo_dicts()

# =============================================================================
# 作業主流程
# =============================================================================
def process_sheet(sheet, program: Program, xls_cell: ExcelCell):
    """處理整個工作表"""
    config = program
    line_start_row = config.line_start_row
    start_row = line_start_row + 2  # 調整為實際起始列
    end_row = start_row + (config.TOTAL_LINES * config.ROWS_PER_LINE)
    rows_per_line = config.ROWS_PER_LINE
    total_lines = config.TOTAL_LINES
    start_col = config.start_col
    end_col = config.end_col
    han_ji_row_offset = config.han_ji_row_offset

    #--------------------------------------------------------------------------
    # 處理所有的儲存格
    active_cell = sheet.range(f'{xw.utils.col_name(start_col)}{line_start_row}')
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    #--------------------------------------------------------------------------
    # 處理作用中列(row)的所有儲存格
    #--------------------------------------------------------------------------
    for r in range(1, total_lines + 1):
        if is_eof: break
        line_no = r
        print('=' * 80)
        print(f"處理第 {line_no} 行...")
        row = line_start_row + (r - 1) * rows_per_line + han_ji_row_offset
        new_line = False
        #----------------------------------------------------------------------
        # 處理列中所有欄(col)儲存格
        #----------------------------------------------------------------------
        for c in range(start_col, end_col + 1):
            if is_eof: break
            row = row
            col = c
            active_cell = sheet.range((row, col))
            active_cell.select()

            #------------------------------------------------------------------
            # 處理儲存格
            #------------------------------------------------------------------
            # status_code:
            # 0 = 儲存格內容為：漢字
            # 1 = 儲存格內容為：文字終結符號
            # 2 = 儲存格內容為：換行符號
            # 3 = 儲存格內容為：空白、標點符號等非漢字字元
            print('-' * 80)
            print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")
            # is_eof, new_line = xls_cell._process_cell(active_cell, row, col)
            status_code = xls_cell._process_cell(active_cell, row, col)

            # 檢查是否需因：換行、文章終結，而跳出內層迴圈
            if new_line: break
            if is_eof: break

    # 將字庫 dict 回存 Excel 工作表
    xls_cell.save_all_piau_im_ji_khoo_dicts()


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
        program = Program(wb, args, hanji_piau_im_sheet_name='漢字注音')

        # 建立儲存格處理器
        if args.new:
            # 建立新的字庫工作表
            xls_cell = ExcelCell(
                program=program,
                new_jin_kang_piau_im_ji_khoo_sheet=True,
                new_piau_im_ji_khoo_sheet=True,
                new_khuat_ji_piau_sheet=True,
            )
        else:
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
        # process_sheet(
        #     sheet=sheet,
        #     program=program,
        #     xls_cell=xls_cell,
        # )
        xls_cell._process_sheet(
            sheet=sheet,
        )

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
        logging_exc_error(msg="無法取得 Excel 活頁簿！", error=None)
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
        try:
            wb.save()
            file_path = wb.fullname
            logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案異常！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
        return EXIT_CODE_SUCCESS

    # =========================================================================
    # (5) 結束程式
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
    import os

    from dotenv import load_dotenv
    load_dotenv()
    DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
    #============================================================================
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