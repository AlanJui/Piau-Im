"""
a200_查找及填入漢字標音.py v0.2.2.6

將【漢字注音】工作表中的【漢字】欄位，依據【人工標音】或【台語音標】查找
【台語音標】，並填入【台語音標】儲存格及【漢字標音】儲存格。
"""

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
# def _show_separtor_line(source_sheet_name: str, target_sheet_name: str):
#     print('\n\n')
#     print("=" * 100)
#     print(f"使用【{source_sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
#     print("=" * 100)


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

        # --------------------------------------------------------------------------
        # 處理作用中列(row)的所有儲存格
        # --------------------------------------------------------------------------
        active_cell = sheet.range(f"{xw.utils.col_name(start_col)}{line_start_row}")
        active_cell.select()

        is_eof = False
        for line_no in range(1, total_lines + 1):
            # 檢查是否到達結尾
            if is_eof or line_no > total_lines:
                break

            # 顯示目前處理【第 n 行】
            self._show_separtor_line(f"處理第 {line_no} 行...")

            # 調整 row 值至【漢字】儲存格所在列
            # （每【行（line）】由 4【列（row）】所構成，漢字在第 3 列：5, 9, 13, ... ）
            row = line_start_row + (line_no - 1) * rows_per_line + han_ji_row_offset

            # ----------------------------------------------------------------------
            # 處理列中所有欄(col)儲存格
            # ----------------------------------------------------------------------
            for c in range(start_col, end_col + 1):
                # 初始化每列所需使用變數
                status_code = 0

                # 將目前處理之儲存格，設為作用中儲存格
                row = row
                col = c
                active_cell = sheet.range((row, col))
                active_cell.select()

                # 顯示正要處理的儲存格座標位置
                print("-" * 80)
                print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")

                # ------------------------------------------------------------------
                # 處理儲存格
                # ------------------------------------------------------------------
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
                    break

        # 將字庫 dict 回存 Excel 工作表
        self.save_all_piau_im_ji_khoo_dicts()


def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件

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
        if args.new:
            # 建立新的字庫工作表
            xls_cell = CellProcessor(
                program=program,
                new_jin_kang_piau_im_ji_khoo_sheet=True,
                new_piau_im_ji_khoo_sheet=True,
                new_khuat_ji_piau_sheet=True,
            )
        else:
            xls_cell = CellProcessor(
                program=program,
                new_jin_kang_piau_im_ji_khoo_sheet=False,
                new_piau_im_ji_khoo_sheet=False,
                new_khuat_ji_piau_sheet=False,
            )

        # --------------------------------------------------------------------------
        # 處理作業開始
        # --------------------------------------------------------------------------
        # 處理工作表
        sheet_name = "漢字注音"
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 處理整張工作表的各個儲存格
        xls_cell._process_sheet(
            sheet=sheet,
        )

        # --------------------------------------------------------------------------
        # 處理作業結束
        # --------------------------------------------------------------------------
        print("=" * 80)
        logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging_exception(
            msg=f"在【{sheet_name}】工作表，自動為【漢字】查找【台語音標】作業，發生例外！",
            error=e,
        )
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

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        logging_exc_error(msg="無法取得 Excel 活頁簿！", error=None)
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        exit_code = process(wb, args)
    except Exception as e:
        msg = f"作業程序發生異常，終止執行：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"處理作業發生異常，終止程式執行：{program_name}（處理作業程序，返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 儲存檔案
        if not Program.save_workbook_as_new_file(wb=wb):
            return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束程式
    # =========================================================================
    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 測試程式
# =============================================================================
def test_01():
    """測試 HanJiTian 類別"""
    import os

    from dotenv import load_dotenv

    load_dotenv()
    DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
    # ============================================================================
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
                    print(
                        f"  台語音標：{item['台語音標']}, 常用度：{item.get('常用度', 'N/A')}, 說明：{item.get('摘要說明', 'N/A')}"
                    )
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
        description="依【漢字】查找【台語音標】並轉換成【漢字標音】",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python a200_查找及填入漢字標音.py          # 執行一般模式
  python a200_查找及填入漢字標音.py -new     # 建立新的字庫工作表
  python a200_查找及填入漢字標音.py -test    # 執行測試模式
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )
    parser.add_argument(
        "--new",
        action="store_true",
        help="建立新的字庫工作表",
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
