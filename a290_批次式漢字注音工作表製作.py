"""a290_批次式漢字注音工作表製作.py

【漢字注音】工作表中，各【漢字】標注之【漢字標音】，以【批次作業方式】變更。

更新紀錄：
 - v0.2.6 2024-06-17:
    變更程式架構，改成套用類別 CellProcessor，借助物件導向程式之【繼承】與【覆蓋】方法，
    以實現【批次式漢字注音工作表製作】功能。
"""

import logging
import sys
from pathlib import Path

import xlwings as xw

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

        # 處理所有的儲存格
        # active_cell = sheet.range((config.line_start_row, config.start_col))
        active_cell = sheet.range(
            f"{xw.utils.col_name(self.program.start_col)}{self.program.line_start_row}"
        )
        active_cell.select()

        # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
        is_eof = False
        for r in range(1, self.program.TOTAL_LINES + 1):
            if is_eof:
                break
            line_no = r

            # 顯示【作用儲存格】位置
            print("-" * 60)
            print(f"處理第 {line_no} 行...")
            row = (
                self.program.line_start_row
                + (r - 1) * self.program.ROWS_PER_LINE
                + self.program.han_ji_row_offset
            )

            new_line = False
            for c in range(self.program.start_col, self.program.end_col + 1):
                if is_eof:
                    break
                if new_line:
                    break  # 跳出內層迴圈，進入下一行處理
                row = row
                col = c
                active_cell = sheet.range((row, col))
                active_cell.select()

                # 顯示正要處理的儲存格座標位置
                print("-" * 60)
                print(f"儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）")

                # ------------------------------------------------------------------
                # 處理儲存格
                # ------------------------------------------------------------------
                # status_code:
                # 0 = 儲存格內容為：漢字
                # 1 = 儲存格內容為：文字終結符號
                # 2 = 儲存格內容為：換行符號
                # 3 = 儲存格內容為：空白、標點符號等非漢字字元
                status_code = 0
                status_code = self._process_cell(active_cell)

                # 檢查是否需因：換行、文章終結，而跳出內層迴圈
                if status_code == 1:
                    is_eof = True
                    break
                elif status_code == 2:
                    new_line = True
                    break

            # 將字庫 dict 回存 Excel 工作表
            # self.save_all_piau_im_ji_khoo_dicts()


def process(wb, args) -> int:
    """
    為【漢字】之【漢字標音】，以批次作業方式，完成各種標音方法標注。

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
        # 初始化 process config
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器
        if args.new:
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
    except Exception as e:
        logging_exc_error(msg="初始化作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # ------------------------------------------------------------------------------
    # 處理作業
    # ------------------------------------------------------------------------------
    try:
        # 待製作之【工作表清單】
        piau_im_name_list = [
            # '台語音標',
            # "雅俗通",
            "十五音",
            # "閩拼調號",
            "閩拼調符",
            "方音符號",
            "台羅拼音",
            "白話字",
        ]

        # 處理工作表
        for piau_im_name in piau_im_name_list:
            print("=" * 80)
            print(f"處理【標音方法】：{piau_im_name} ...")
            # 設定目前使用的標音方法
            program.piau_im_huat = piau_im_name

            # 切換工作表
            sheet_name = f"漢字注音【{piau_im_name}】"
            # 使用【漢字注音】工作表複製新工作表
            try:
                if sheet_name not in [sheet.name for sheet in wb.sheets]:
                    # 複製工作表
                    source_sheet = wb.sheets["漢字注音"]
                    new_sheet = source_sheet.copy(name=sheet_name, after=source_sheet)
                    print(f"✅ 已複製【漢字注音】工作表為 '{sheet_name}'")
                else:
                    print(f"⚠️ 工作表 '{sheet_name}' 已存在")

            except Exception as e:
                raise ValueError(f"無法找到或建立工作表 '{sheet_name}'：{e}")

            try:
                # sheet = wb.sheets[sheet_name]
                # sheet.activate()

                # 處理整張工作表的各個儲存格
                # xls_cell._process_sheet(sheet)
                new_sheet.activate()
                xls_cell._process_sheet(new_sheet)
            except Exception as e:
                logging_exception(
                    msg=f"在【{sheet_name}】工作表，為【漢字】標注{program.piau_im_huat}【漢字標音】作業，發生例外！",
                    error=e,
                )
                raise
    except Exception as e:
        logging_exception(
            msg=f"程式：{program.program_name} ，執行時發生異常問題！",
            error=e,
        )
        raise

    # ------------------------------------------------------------------------------
    # 處理作業結束
    # ------------------------------------------------------------------------------
    print("=" * 80)
    logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    """主程式"""
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
    try:
        # 取得 Excel 活頁簿
        wb = None
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
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
    pass


if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="輸入參數說明：\n  - --test: 執行測試模式\n  - --new: 建立新的字庫工作表",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python ao00_xyz.py            # 執行一般模式
  python ao00_xyz.py -new       # 建立新的字庫工作表
  python ao00_xyz.py -test      # 執行測試模式
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
    # args.program_name = Path(__file__).stem

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
