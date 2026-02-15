#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
簡單說明作業流程如下：
遇【作用儲存格】填入【引用既有的漢字標音】符號（【=】）時，漢字的【台語音標】
自【人工標音字庫】工作表查找，並轉換成【漢字標音】。

在【漢字注音】工作表，若使用者曾對某漢字以【人工標音】儲存格手動標音過，則再
次遇到相同之漢字時，若在【人工標音】儲存格填入【=】符號（表示引用既有的標音），
則使用者可省去重新標音的麻煩；而程式會負責自【人工標音字庫】工作表查找該漢字的
【台語音標】，並轉換成【漢字標音】填入對應的儲存格。

顧及使用者可能會有記憶錯誤的狀況發生，若在【人工標音字庫】工作表找不到對應的
【台語音標】時，程式會再自【標音字庫】工作表查找一次，若仍找不到，則將該漢字
記錄到【缺字表】工作表，以便後續處理。
"""
# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# 載入自訂模組
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
    logging_exception,  # noqa: F401
    logging_process_step,  # noqa: F401
    logging_warning,  # noqa: F401
)
from mod_標音 import is_han_ji
from mod_程式 import ExcelCell, Program

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
# 資料類別：儲存處理配置
# =========================================================================
class CellProcessor(ExcelCell):
    """
    個人字典查詢專用的儲存格處理器
    繼承自 ExcelCell
    覆蓋以下方法以實現個人字典查詢功能：
    - _process_han_ji(): 使用個人字典查詢漢字讀音
    - process_cell(): 處理單一儲存格
    - _process_sheet(): 處理整個工作表
    """

    def __init__(
        self,
        program: Program,
        new_jin_kang_piau_im_ji_khoo_sheet: bool = False,
        new_piau_im_ji_khoo_sheet: bool = False,
        new_khuat_ji_piau_sheet: bool = False,
    ):
        # 調用父類別（MengDianExcelCell）的建構子
        super().__init__(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=new_jin_kang_piau_im_ji_khoo_sheet,
            new_piau_im_ji_khoo_sheet=new_piau_im_ji_khoo_sheet,
            new_khuat_ji_piau_sheet=new_khuat_ji_piau_sheet,
        )

    def _process_cell(
        self,
        cell,
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
        row = cell.row  # 取得【漢字】儲存格的列號
        col = cell.column  # 取得【漢字】儲存格的欄號

        cell_value = cell.value
        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        tai_gi_im_piau = cell.offset(-1, 0).value  # 台語音標
        han_ji_piau_im = cell.offset(1, 1).value  # 漢字標音

        # 初始化樣式
        self._reset_cell_style(cell)

        # 確保 cell_value 務必是【漢字】，故需篩飾【特殊字元】
        if cell_value == "φ":
            # 【文字終結】
            print("【文字終結】")
            return 1  # 文章終結符號
        elif cell_value == "\n":
            # 【換行】
            print("【換行】")
            return 2  # 【換行】
        elif cell_value is None or str(cell_value).strip() == "":
            print("【空白】")
            return 3  # 空白或標點符號
        elif not is_han_ji(cell_value):
            # 處理【標點符號】、【英數字元】、【其他字元】
            msg = self._process_non_han_ji(cell)
            print(msg)
            return 3  # 空白或標點符號

        # ======================================================================
        # 自此以下，儲存格存放【漢字】。每個【漢字】儲存格有三種可能：
        # 1. 【無標音漢字】：在【個人字典】找不到讀音，故【台語音標】、【漢字標音】
        #     儲存格為空白。在【缺字表】工作表有紀錄登錄；
        # 2. 【自動標音漢字】：在【個人字典】找到讀音，故【台語音標】、【漢字標音】
        #     儲存格已有讀音標注。在【標音字庫】有紀錄登錄；
        # 3. 【人工標音漢字】：在【人工標音】儲存格，有手動輸入之【台羅拼音】、【TLPA音標】
        #     。或是【=】（引用【人工標音】）。在【人工標音字庫】有紀錄登錄。
        # ======================================================================

        # 檢查是否為【無標音漢字】
        if (
            not tai_gi_im_piau
            or str(tai_gi_im_piau).strip() == ""
            and not han_ji_piau_im
        ):
            self._process_bo_thok_im(cell)
            return 0  # 漢字

        # 檢查是否為【人工標音漢字】
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            self._show_msg(row, col, cell_value)
            self._process_jin_kang_piau_im(cell)
            return 0  # 漢字

        # 處理【自動標音漢字】
        self._process_han_ji(cell)
        return 0  # 漢字

    def _process_sheet(self, sheet):
        """處理整個工作表"""
        program = self.program

        # 自【作用儲存格】取得【Excel 儲存格座標】(列,欄) 座標
        try:
            active_cell = sheet.api.Application.ActiveCell
            # 顯示【作用儲存格】位置
            active_row = active_cell.Row
            active_col = active_cell.Column
            active_col_name = xw.utils.col_name(active_col)
        except Exception:
            raise ValueError("無法取得作用儲存格")

        # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
        line_start_row = (
            self.program.line_start_row
        )  # 第一行【標音儲存格】所在 Excel 列號: 3
        line_no = ((active_row - line_start_row + 1) // self.program.ROWS_PER_LINE) + 1
        han_ji_row = (line_no * program.ROWS_PER_LINE) + program.han_ji_row_offset - 1
        col = active_col
        cell = sheet.range((han_ji_row, col))
        # 處理儲存格
        act_cell_addr = f"作用儲存格：{active_col_name}{active_row}（{active_cell.Row}, {active_cell.Column}）"
        han_ji_cell_addr = (
            f"漢字儲存格：{active_col_name}{han_ji_row}（{han_ji_row}, {col}）"
        )
        print(f"{act_cell_addr} ==> {han_ji_cell_addr}")
        self._process_cell(cell)


# =========================================================================
# 主要處理函數
# =========================================================================
def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

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
        # 初始化 Program 配置
        # --------------------------------------------------------------------------
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")

        # 建立萌典專用的儲存格處理器（繼承自 ExcelCell）
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=(
                args.new if hasattr(args, "new") else False
            ),
            new_piau_im_ji_khoo_sheet=args.new if hasattr(args, "new") else False,
            new_khuat_ji_piau_sheet=args.new if hasattr(args, "new") else False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 作業處理中
    # --------------------------------------------------------------------------
    try:
        # 處理工作表
        sheet_name = program.hanji_piau_im_sheet_name
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        xls_cell._process_sheet(sheet=sheet)

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 處理作業結束
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main(args):
    # =========================================================================
    # 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    program_name = current_file_path.stem

    # =========================================================================
    # 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    """主程式 - 從 Excel 呼叫或直接執行"""
    try:
        # 取得 Excel 活頁簿
        wb = None
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
            return EXIT_CODE_NO_FILE

        if not wb:
            logging.error("無法取得 Excel 活頁簿")
            return EXIT_CODE_NO_FILE

        # ==================================================================
        # 執行處理作業
        # ==================================================================
        print("=" * 80)
        print("無限循環模式：請在 Excel 中選擇任一儲存格後按 Enter 查詢")
        print("按 Ctrl+C 終止程式")
        print("=" * 80)
        sheet_name = "漢字注音"

        # 無限循環
        while True:
            try:
                # 等待使用者按 Enter
                input(
                    "\n請在 Excel 選擇【作用儲存格】後按 Enter 繼續（Ctrl+C 終止）..."
                )

                # 確保工作表為作用中
                wb.sheets[sheet_name].activate()

                exit_code = process(wb=wb, args=args)
                if exit_code != EXIT_CODE_SUCCESS:
                    print(f"⚠️  處理結果：exit_code = {exit_code}")

            except KeyboardInterrupt:
                print("\n\n使用者中斷程式（Ctrl+C）")
                print("=" * 70)
                # ==================================================================
                # 儲存檔案
                # ==================================================================
                if exit_code == EXIT_CODE_SUCCESS:
                    try:
                        wb.save()
                        file_path = wb.fullname
                        logging_process_step(f"儲存檔案至路徑：{file_path}")
                    except Exception as e:
                        logging_exc_error(msg="儲存檔案異常！", error=e)
                        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案
                    return EXIT_CODE_SUCCESS

            except Exception as e:
                logging.error(f"處理錯誤：{e}")
                print(f"❌ 錯誤：{e}")
                # 發生錯誤時繼續循環，不中斷程式
                continue

    except KeyboardInterrupt:
        print("\n\n使用者中斷程式（Ctrl+C）")
        print("=" * 70)
        return EXIT_CODE_SUCCESS
    except Exception as e:
        logging.exception(f"程式執行失敗: {e}")
        return EXIT_CODE_UNKNOWN_ERROR


def test_han_ji_tian():
    """測試 HanJiTian 類別"""
    # =========================================================================
    # 載入環境變數
    # =========================================================================
    import os

    from dotenv import load_dotenv

    from mod_ca_ji_tian import HanJiTian  # 新的查字典模組

    # 預設檔案名稱從環境變數讀取
    load_dotenv()
    DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")

    print("=" * 80)
    print("測試 HanJiTian 查詢功能")
    print("=" * 80)

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
                print("  查無資料")

        print("\n" + "=" * 80)
        print("測試完成")
        print("=" * 80)

    except Exception as e:
        print(f"測試失敗：{e}")
        import traceback

        traceback.print_exc()


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
    parser.add_argument(
        "--new",
        action="store_true",
        help="建立新的標音字庫工作表",
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
