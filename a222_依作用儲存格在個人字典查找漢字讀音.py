# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
from pathlib import Path
from typing import Tuple

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
from mod_帶調符音標 import is_han_ji
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
# 自訂 ExcelCell 子類別：覆蓋特定方法以實現萌典查詢功能
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

    # =================================================================
    # 輔助方法
    # =================================================================
    def check_coordinate_exists(self, row: int, col: int, coord_list: list) -> bool:
        """
        檢查座標是否存在於座標列表中

        Args:
            row: 列號
            col: 欄號
            coord_list: 座標列表

        Returns:
            bool: 座標是否存在
        """
        if not coord_list:
            return False
        return (row, col) in coord_list

    # =================================================================
    # 覆蓋父類別的方法
    # =================================================================
    def _process_han_ji(
        self,
        han_ji: str,
        cell,
        row: int,
        col: int,
    ) -> Tuple[str, bool]:
        """
        處理漢字 - 使用萌典 API 查詢讀音
        ⚠️ 覆蓋父類別的方法 - 使用萌典而非本地資料庫

        Args:
            han_ji: 要查詢的漢字
            cell: Excel 儲存格物件
            row: 儲存格列號
            col: 儲存格欄號

        Returns:
            (message, success): 處理訊息和是否成功
        """
        if han_ji == "":
            return "【空白】", False

        result = self.program.ji_tian.han_ji_ca_piau_im(
            han_ji=han_ji,
            ue_im_lui_piat=self.program.ue_im_lui_piat,
        )

        # 查無此字
        if not result:
            # 記錄到缺字表
            self.khuat_ji_piau_ji_khoo_dict.add_entry(
                han_ji=han_ji,
                tai_gi_im_piau="",
                hau_ziann_im_piau="N/A",
                coordinates=(row, col),
            )
            return f"【{han_ji}】查無此字！", False

        # 有多個讀音
        print(
            f"漢字儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）：【{han_ji}】有 {len(result)} 個讀音：{result}"
        )

        # 顯示所有讀音選項
        piau_im_options = []
        for idx, tai_lo_ping_im in enumerate(result):
            # 轉換音標
            tai_gi_im_piau, han_ji_piau_im = self._convert_piau_im_by_entry(
                tai_lo_ping_im
            )
            piau_im_options.append((tai_gi_im_piau, han_ji_piau_im))
            msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
            print(f"{idx + 1}. {msg}")

        # 讓使用者選擇讀音
        user_input = input(
            "\n請選擇讀音編號（直接按 Enter 略過，輸入編號後按 Enter 填入）："
        ).strip()

        if user_input == "":
            # 只瀏覽，不填入
            print("略過填入")
            return f"【{han_ji}】已顯示 {len(result)} 個讀音（未填入）", False

        try:
            choice = int(user_input)
            if 1 <= choice <= len(result):
                # 填入選擇的讀音
                tai_gi_im_piau, han_ji_piau_im = piau_im_options[choice - 1]
                cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音儲存格
                cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標儲存格
                cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音儲存格

                # 在【人工標音字庫】增添【該字】指向【漢字注音】之【座標】紀錄
                self.jin_kang_piau_im_ji_khoo_dict.add_or_update_entry_by_coordinate(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    hau_ziann_im_piau="N/A",
                    coordinates=(row, col),
                )
                # 儲存更新後的【人工標音字庫】至工作表
                self.jin_kang_piau_im_ji_khoo_dict.save_to_sheet(
                    wb=self.program.wb,
                    sheet_name=self.jin_kang_piau_im_ji_khoo_dict.name,
                )

                # 自【標音字庫】移除【該字】指向【漢字注音】之【座標】紀錄
                row_no = self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
                    han_ji=han_ji, coordinate=(row, col)
                )
                if row_no != -1:
                    _, entry = self.piau_im_ji_khoo_dict.get_entry_by_row_no(row_no)
                    exist = self.check_coordinate_exists(
                        row=row,
                        col=col,
                        coord_list=entry["coordinates"],
                    )
                    if exist:
                        self.piau_im_ji_khoo_dict.remove_coordinate(
                            han_ji=han_ji,
                            coordinate=(row, col),
                            entry_to_delete_if_empty=False,
                        )
                        # 儲存更新後的【標音字庫】至工作表
                        self.piau_im_ji_khoo_dict.save_to_sheet(
                            wb=self.program.wb,
                            sheet_name=self.piau_im_ji_khoo_dict.name,
                        )

                print(
                    f"已填入第 {choice} 個讀音：[{tai_gi_im_piau}] /【{han_ji_piau_im}】"
                )
                return f"【{han_ji}】已填入第 {choice} 個讀音", True
            else:
                print(f"無效的選擇：{choice}（超出範圍）")
                return f"【{han_ji}】選擇無效", False
        except ValueError:
            print(f"無效的輸入：{user_input}")
            return f"【{han_ji}】輸入無效", False

    def _process_cell(
        self,
        cell,
        row: int,
        col: int,
    ) -> bool:
        """
        處理單一儲存格

        Returns:
            is_eof: 是否已達文件結尾
            new_line: 是否需換行
        """
        # 初始化樣式
        self._reset_cell_style(cell)

        cell_value = cell.value

        # 確保 cell_value 務必是【漢字】，故需篩飾【特殊字元】
        if cell_value == "φ":
            # 【文字終結】
            print(f"【{cell_value}】：【文章結束】結束行處理作業。")
            return True, True
        elif cell_value == "\n":
            # 【換行】
            print("【換行】：結束行中各欄處理作業。")
            return False, True
        elif not is_han_ji(cell_value):
            # 處理【標點符號】、【英數字元】、【其他字元】
            self._process_non_han_ji(cell_value)
            return False, False
        else:
            self._process_han_ji(cell_value, cell, row, col)
            return False, False

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
            print(
                f"作用儲存格：{active_col_name}{active_row}（{active_cell.Row}, {active_cell.Column}）"
            )
        except Exception:
            raise ValueError("無法取得作用儲存格")

        # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
        line_start_row = (
            self.program.line_start_row
        )  # 第一行【標音儲存格】所在 Excel 列號: 3
        line_no = ((active_row - line_start_row + 1) // self.program.ROWS_PER_LINE) + 1
        row = (line_no * program.ROWS_PER_LINE) + program.han_ji_row_offset - 1
        col = active_col
        cell = sheet.range((row, col))
        # 處理儲存格
        self._process_cell(cell, row, col)


# =========================================================================
# 主要處理函數
# =========================================================================
def process(wb, args) -> int:
    """
    查詢漢字讀音並標注 - 使用【個人字典】

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
