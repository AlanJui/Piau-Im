"""
a221_依字庫工作表之作用儲存格在個人字典查找漢字讀音.py v0.0.1
功能說明：
    在【缺字表/標音字庫】等字庫工作表的【作用儲存格】位置，在個人字典中查找漢字讀音。
更新紀錄：
v0.0.1: 自 a222.py 改寫。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import re
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# 載入自訂模組
from mod_excel_access import excel_address_to_row_col, get_active_cell_address, get_line_no_by_row, get_row_by_line_no
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
    logging_process_step,  # noqa: F401
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
    def _za_ji_tain_au_thiam_jin_kang_piau_im(self, active_cell):
        """查字典後填入工標音"""
        tai_gi_im_piau = ""

        # 依據【作用儲存格】之【漢字】，從【自用字典】查詢【台語音標】
        tai_gi_im_piau = self._han_ji_ca_piau_im_kap_cu_tik(active_cell)
        if tai_gi_im_piau is None:
            return None, None

        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = self.convert_tai_gi_im_piau_to_han_ji_piau_im(tai_gi_im_piau=tai_gi_im_piau)

        active_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        active_cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音

        return tai_gi_im_piau, han_ji_piau_im

    def _convert_coord_list_str_to_tuples(self, coordinates_str):
        # 利用正規表達式，解析所有形如 (row, col) 之座標
        coordinate_tuples = []
        coordinate_tuples = re.findall(r"\((\d+)\s*,\s*(\d+)\)", str(coordinates_str))
        return coordinate_tuples

    def _update_piau_im_ji_khoo_from_khuat_ji_piau(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        coordinates: list[tuple],
    ) -> None:
        """
        將原本在【缺字表】工作表之【資料紀錄】，寫入【標音字庫】工作表，以示該漢字已查得
        【台語音標】，並已在【漢字注音】工作表完成【台語音標】及【漢字標音】之標注工作。

        處理步驟：
        （1）依【座標清單】，將【漢字＋台語音標＋座標】逐筆登錄至【標音字庫】字典；
        （2）自【缺字表】字典，逐一移除該漢字之【座標】（座標清空時，整筆紀錄一併移除）；
        （3）將上述兩字典之內容，回寫至【標音字庫】與【缺字表】工作表。

        Args:
            han_ji: 已查得讀音之漢字
            tai_gi_im_piau: 自個人字典查得之台語音標
            coordinates: 指向【漢字注音】工作表【漢字】儲存格之座標清單
        """
        wb = self.program.wb

        for coord in coordinates:
            # 座標若為字串（re.findall 解析結果），先轉換成整數座標
            coordinate = (int(coord[0]), int(coord[1]))

            # （1）在【標音字庫】字典，新增（或更新）一筆資料紀錄
            self.piau_im_ji_khoo_dict.add_or_update_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                hau_ziann_im_piau="N/A",
                coordinates=coordinate,
            )

            # （2）自【缺字表】字典，移除該漢字之此一【座標】
            self.khuat_ji_piau_ji_khoo_dict.remove_coordinate_by_han_ji_and_coordinate(
                han_ji=han_ji,
                coordinate=coordinate,
            )

        # （3）將兩字典之內容，回寫至各自之工作表
        self.piau_im_ji_khoo_dict.save_to_sheet(
            wb=wb, sheet_name=self.piau_im_ji_khoo_dict.name
        )
        self.khuat_ji_piau_ji_khoo_dict.save_to_sheet(
            wb=wb, sheet_name=self.khuat_ji_piau_ji_khoo_dict.name
        )

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
        # 指定【漢字注音】工作表為【作用工作表】
        source_sheet_name = "缺字表"
        source_sheet = wb.sheets[source_sheet_name]
        source_sheet.activate()

        # 取得【作用儲存格】
        active_cell_address = get_active_cell_address()
        active_cell = source_sheet.range(active_cell_address)
        row, col = excel_address_to_row_col(active_cell_address)
        # 自作用儲存格所在列，取得【漢字】（在A欄）
        han_ji = source_sheet.range((row, 1)).value
        tai_gi_im_piau = source_sheet.range((row, 2)).value

        # ----------------------------------------------------------------------
        # 依【作用儲存格】所在列（Row），自A欄取得【漢字】，B欄取得【台語音標】。
        # 並依上述資料判斷，是否程式還需繼續。
        # ----------------------------------------------------------------------
        if not is_han_ji(han_ji):
            msg=f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_SUCCESS
        elif not tai_gi_im_piau == "N/A":
            msg=f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，台語音標為【{tai_gi_im_piau}】，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_SUCCESS

        # ----------------------------------------------------------------------
        # 依取得之【漢字】，在漢字庫資料庫尋找【台語音標】
        # ----------------------------------------------------------------------
        cell = source_sheet.range((row, 1))
        tai_gi_im_piau = xls_cell._han_ji_ca_piau_im_kap_cu_tik(cell)
        # 依據使用者輸入之【台語音標】轉換為【漢字標音】
        han_ji_piau_im = xls_cell.convert_tai_gi_im_piau_to_han_ji_piau_im(
            tai_gi_im_piau=tai_gi_im_piau,
        )
        if not tai_gi_im_piau or not han_ji_piau_im:
            print(f"因異常狀況，終止作業......")
            print(f"{active_cell_address} 漢字：【{han_ji}】，台語音標：【{tai_gi_im_piau}】，漢字標音：【{han_ji_piau_im}】。")
            return EXIT_CODE_PROCESS_FAILURE

        # 將字典查得之【台語音標】填入【缺字表】工作表
        source_sheet.range((row, 2)).value = tai_gi_im_piau

        # ----------------------------------------------------------------------
        # 將已查得之【台語音標】，及轉換所得之【漢字標音】，依據【缺字表】工作表在【座標】欄
        # 取得之【座標清單】，一一回填【漢字注音】工作表
        # ----------------------------------------------------------------------
        # 取得【座標】欄之【座標清單】：內容為字串，格式如 "(5, 17); (33, 8)"
        target_sheet = wb.sheets[program.hanji_piau_im_sheet_name]
        coordinates_str = source_sheet.range((row, 4)).value
        if coordinates_str:
            # 利用正規表達式，解析所有形如 (row, col) 之座標
            coordinate_tuples =  xls_cell._convert_coord_list_str_to_tuples(coordinates_str)
            for tup in coordinate_tuples:
                coord_row, coord_col = int(tup[0]), int(tup[1])
                # 將【台語音標】和【漢字標音】寫入【漢字注音】工作表：
                # 【台語音標】在【漢字】儲存格上一列；【漢字標音】在下一列
                target_sheet.range((coord_row - 1, coord_col)).value = tai_gi_im_piau
                target_sheet.range((coord_row + 1, coord_col)).value = han_ji_piau_im

                # 顯示已完成的工作結果
                print(f"📌 ({coord_row}, {coord_col}) 漢字：{han_ji}，台語音標：{tai_gi_im_piau}")

        # -------------------------------------------------------------------------
        # 將原本在【缺字表】工作表之【資料紀錄】，寫入【標音字庫】工作表，以示該漢字已查得
        # 【台語音標】，並已在【漢字注音】工作表完成【漢字】的【台語音標】及【漢字注音】完成標注
        # 工作。
        # -------------------------------------------------------------------------
        xls_cell._update_piau_im_ji_khoo_from_khuat_ji_piau(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            coordinates=coordinate_tuples,
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
        # -------------------------------------------------------------------------
        # 因【標音字庫】依【漢字】之【座標】紀錄，更新【漢字注音】工作表中對映之【台語音標】及【漢字標音】；導致
        # 【作用儲存格】之 Excel Address 已變更，需將之校正回歸。
        # -------------------------------------------------------------------------
        source_sheet.activate()
        active_cell.select()  # 選取【作用儲存格】，以確保游標位置正確

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
