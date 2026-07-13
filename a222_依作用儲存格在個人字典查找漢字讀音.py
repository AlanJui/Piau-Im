"""
a222_依作用儲存格在個人字典查找漢字讀音.py V0.3
功能說明：
    依作用儲存格位置，在個人字典中查找漢字讀音。
更新紀錄：
V0.2.6 2026-02-08: 修正顯示漢字讀音選項時的輸出格式；加入【常用度】欄位，
    便於使用者選擇適合的讀音。
V0.2.7 2026-02-09: 變更【個人字典查找漢字標音作業】，查得之【台語音標】不
    記錄於【人工標音字庫】工作表中；而是用於更換【標音字庫】工作表紀錄；另外，
    自【個人字典】查得之【台語音標】，依【座標】欄已登載之各【座標】，更新
    【漢字注音】工作表中，對應【漢字】之【台語音標】及【漢字標音】。
v0.2.8 2026-02-09: 查得漢字之標音並選用後，因更新【標音字庫】工作表中對應
    【漢字】之【台語音標】及【漢字標音】，導致【作用儲存格】之 Excel Address
    已變更，需將之校正回歸。
v0.2.9 2026-02-13: 修正原【無標音漢字】與【缺字表】工作表無法正常運作之問題。
v0.2.12 2026-03-18: 改善 _bo_thok_im() 方法，當【台語音標】或【漢字標音】為空值時，
    均屬標音異常，很可能起因於字典當無該漢字之讀音資料，或其它原因，故要求使用者重新輸入。
v0.2.13 2026-03-21: 修正查字典時，顯示所有讀音的預設值為 True。
v0.2.14 2026-03-22: 仿 a260_依字典查得結果填入人工標音.py 修改及優化。
v0.3 2026-07-10: 自 a250.py 修改而得。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from mod_excel_access import (
    excel_address_to_row_col,
    get_active_cell_address,
    get_line_no_by_row,
    get_row_by_line_no,
)
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)

# 載入自訂模組/函式
# from mod_標音 import convert_tl_with_tiau_hu_to_tlpa, is_han_ji, kam_si_u_tiau_hu, tlpa_tng_han_ji_piau_im
from mod_標音 import is_han_ji, tlpa_tng_han_ji_piau_im
from mod_程式 import ExcelCell, Program

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

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

    # =================================================================
    # 覆蓋父類別的方法
    # =================================================================
    # def check_coordinate_exists(self, row: int, col: int, coord_list: list) -> bool:
    #     """
    #     檢查座標是否存在於座標列表中

    #     Args:
    #         row: 列號
    #         col: 欄號
    #         coord_list: 座標列表

    #     Returns:
    #         bool: 座標是否存在
    #     """
    #     if not coord_list:
    #         return False
    #     return (row, col) in coord_list

    def _update_piau_im_ji_khoo_worksheet(
        self,
        row: int,
        col: int,
        han_ji: str,
        tai_gi_im_piau: str,
        han_ji_piau_im: str,
    ) -> None:
        """
        訂正作業步驟 5–7：更新【標音字庫】，並同步【漢字注音】工作表。

        5. 依【漢字】與作用儲存格【座標】，在【標音字庫】找到既有資料紀錄
           （該紀錄仍保留訂正前之【台語音標】）。
        6. 將該紀錄之【台語音標】更新為使用者選定／輸入之新音標。
        7. 取出該紀錄【座標】欄之【座標清單】，逐一更新【漢字注音】工作表
           對應儲存格之【台語音標】與【漢字標音】。
        """
        # row = cell.row  # 漢字儲存格所在之 Row
        # col = cell.column
        # han_ji = cell.value
        # # 使用者已選定／輸入之新【台語音標】與轉換後之【漢字標音】
        # tai_gi_im_piau = cell.offset(-1, 0).value
        # han_ji_piau_im = cell.offset(1, 0).value

        # ------------------------------------------------------------------
        # 步驟 5：依【漢字】與【座標】，在【標音字庫】找到既有資料紀錄
        # ------------------------------------------------------------------
        row_no = self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
            han_ji=han_ji,
            coordinate=(row, col),
        )
        if row_no == -1:
            print(f"⚠️  【標音字庫】查無【{han_ji}】座標 ({row}, {col}) 之紀錄，略過字庫更新。")
            return

        # ------------------------------------------------------------------
        # 步驟 6–7：更新【台語音標】，並依【座標清單】同步【漢字注音】工作表
        # ------------------------------------------------------------------
        self.update_piau_im_worksheet_entry(
            coordinate=(row, col),
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            han_ji_piau_im=han_ji_piau_im,
            piau_im_ji_khoo_dict=self.piau_im_ji_khoo_dict,
            row_no=row_no,
        )

    def _process_sheet(self):

        try:
            # return super()._process_sheet(sheet, show_cell_address)
            # --------------------------------------------------------------------------
            # 處理作業開始
            # --------------------------------------------------------------------------
            # source_sheet_name = "漢字注音"
            source_sheet_name = self.program.hanji_piau_im_sheet_name
            wb = self.program.wb

            # ----------------------------------------------------------------------
            # 取得【作用儲存格】
            # ----------------------------------------------------------------------
            # 指定【漢字注音】工作表為【作用工作表】
            source_sheet = wb.sheets[source_sheet_name]
            source_sheet.activate()

            active_cell_address = get_active_cell_address()
            active_cell = source_sheet.range(active_cell_address)
            row, col = excel_address_to_row_col(active_cell_address)
            current_line_no = get_line_no_by_row(current_row_no=row)  # 計算行號
            jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = get_row_by_line_no(current_line_no)
            han_ji_cell = source_sheet.range((han_ji_row, col))
            source_sheet.range((han_ji_row, col)).select()  # 選取【漢字】儲存格，以確保游標位置正確
            source_sheet.activate()  # 重新激活工作表以刷新儲存格地址

            # 確認【作用儲存格】為【漢字】
            han_ji = source_sheet.range((han_ji_row, col)).value
            if not is_han_ji(han_ji):
                msg = f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
                print(f">> {msg}")
                return EXIT_CODE_SUCCESS

            # 確認【作用儲存格】的【漢字】有【台語音標】及【漢字標音】，否則可能是字典目前無此【漢字】之讀音資料，
            # 故後續之查字典作業應被略過，直接要求使用者輸入【台語音標】或【台羅拼音】。
            tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
            han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value
            jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
            # 記錄原始的的【人工標音】
            # original_jin_kang_piau_im = jin_kang_piau_im
            original_tai_gi_im_piau = tai_gi_im_piau
            original_han_ji_piau_im = han_ji_piau_im

            if not tai_gi_im_piau or not han_ji_piau_im:
                # ----------------------------------------------------------------------
                # 無法查字典狀況的處理作業：若【漢字注音】工作表，【作用儲存格】之【台語音標】或【漢字標音】
                # 無資料填入，則表【字典】可能無此漢字之【台語音標】。遇此狀況，則【用漢字查字典】之作
                # 業至此終止，改以【人工查尋，手動輸入作業】。
                # ----------------------------------------------------------------------
                # 告知使用者，終止查字作業，需要先人工查字之台語音標，然後手動輸入。
                msg = f"作用儲存格 {active_cell_address} 的漢字【{han_ji}】缺乏【台語音標】或【漢字標音】，可能是字典無此漢字之讀音資料，將略過查字典作業，直接要求使用者輸入【台語音標】或【台羅拼音】。"
                print(f">> {msg}")
                # 取得使用者輸入之【台語音標】或【台羅拼音】
                tai_gi_im_piau = self.get_user_input_piau_im(han_ji=han_ji)
                # 依據使用者輸入之【台語音標】轉換為【漢字標音】
                han_ji_piau_im = self._convert_tai_gi_im_piau_to_han_ji_piau_im(
                    tai_gi_im_piau=tai_gi_im_piau,
                )
                # 將手動輸入之【台語音標】及經程式轉換而得之【漢字標音】，填入【作用儲存格】的【人工標音】
                # 、【台語音標】及【漢字標音】三處儲存格。
                source_sheet.range((tai_gi_im_piau_row, col)).value = tai_gi_im_piau
                source_sheet.range((han_ji_piau_im_row, col)).value = han_ji_piau_im
                source_sheet.range((jin_kang_piau_im_row, col)).value = jin_kang_piau_im
            else:
                # ----------------------------------------------------------------------
                # 使用【漢字】查字典，取得【台語音標】
                # ----------------------------------------------------------------------
                han_ji_position = (han_ji_row, col)
                print(f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}")
                print(f"📌 漢字：{han_ji}")
                print(f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}")

                # 查字典後，將查尋所得之【台語音標】，轉換【漢字標音】，最後回傳
                tai_gi_im_piau, han_ji_piau_im = self._za_ji_tian_au_thiam_jin_kang_piau_im(active_cell=han_ji_cell)
                # 若使用者放棄變更，則結束作業流程
                if not tai_gi_im_piau and not han_ji_piau_im:
                    return EXIT_CODE_SUCCESS

            msg = f"【{han_ji}】[{original_tai_gi_im_piau}] / [{original_han_ji_piau_im}] 變更為： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
            print(f">> 儲存格：{active_cell_address}，{msg}")

            # -------------------------------------------------------------------------
            # 將【台語音標】和【漢字標音】，回寫【漢字注音】工作表之【作用儲存格】
            # -------------------------------------------------------------------------
            self._update_piau_im_ji_khoo_worksheet(
                row=han_ji_row,
                col=col,
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                han_ji_piau_im=han_ji_piau_im,
            )
            # -------------------------------------------------------------------------
            # 在【人工標音字庫】之【字庫表】(dict)，新增該【漢字】之記錄
            # -------------------------------------------------------------------------
            self._update_jin_kang_piau_im_ji_khoo_worksheet_by_move(
                han_ji=han_ji,
                jin_kang_piau_im=tai_gi_im_piau,
                row=han_ji_row,
                col=col,
            )
            # -------------------------------------------------------------------------
            # 更新資料庫中【漢字庫】資料表
            # -------------------------------------------------------------------------
            siong_iong_too_to_use = 0.8 if self.program.ue_im_lui_piat == "文讀音" else 0.6  # 根據語音類型設定常用度
            self.insert_or_update_to_db(
                table_name=self.program.table_name,
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                ue_im_lui_piat=self.program.ue_im_lui_piat,
                siong_iong_too=siong_iong_too_to_use,
            )
            # -------------------------------------------------------------------------
            # 因【標音字庫】依【漢字】之【座標】紀錄，更新【漢字注音】工作表中對映之【台語音標】及【漢字標音】；導致
            # 【作用儲存格】之 Excel Address 已變更，需將之校正回歸。
            # -------------------------------------------------------------------------
            source_sheet.activate()  # 重新激活工作表以刷新儲存格地址
            active_cell.select()  # 選取【作用儲存格】，以確保游標位置正確

            logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
            return EXIT_CODE_SUCCESS
        except Exception as e:
            # 你可以在這裡加上紀錄或處理，例如:
            logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
            # 再次拋出異常，讓外層函式能捕捉
            raise

    # =================================================================
    # 子類別的方法
    # =================================================================

    # =================================================================
    # 輔助方法
    # =================================================================
    def _za_ji_tian_au_thiam_jin_kang_piau_im(self, active_cell):
        """查字典後填入工標音"""
        piau_im_huat = self.program.piau_im_huat
        piau_im = self.program.piau_im
        tai_gi_im_piau = ""

        # 依據【作用儲存格】之【漢字】，從【自用字典】查詢【台語音標】
        tai_gi_im_piau = self._han_ji_ca_piau_im_kap_cu_tik(active_cell)
        if tai_gi_im_piau is None:
            return None, None

        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau,
        )

        return tai_gi_im_piau, han_ji_piau_im


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
    logging_process_step("<=========== 作業開始！==========>")

    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    try:
        # --------------------------------------------------------------------------
        # 初始化 Program 配置
        # --------------------------------------------------------------------------
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器（繼承自 ExcelCell）
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 作業處理中
    # --------------------------------------------------------------------------
    try:
        # 確保工作表為作用中
        sheet_name = program.hanji_piau_im_sheet_name
        sheet = wb.sheets[sheet_name]
        sheet.activate()
        xls_cell._process_sheet()

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 處理作業結束
    # --------------------------------------------------------------------------
    print("\n")
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
    # =========================================================================
    # (0) 程式初始化
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
                input("\n請在 Excel 選擇【作用儲存格】後按 Enter 繼續（Ctrl+C 終止）...")

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
                    print(f"  台語音標：{item['台語音標']}, 常用度：{item.get('常用度', 'N/A')}, 說明：{item.get('摘要說明', 'N/A')}")
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
