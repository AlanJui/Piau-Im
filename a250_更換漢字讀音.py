"""
a250_作用儲存格查找漢字標音.py V0.3

依作用儲存格位置，在個人字典中查找漢字讀音。找到／輸入之漢字讀音，將用
以代換【漢字標音】工作表已登錄之【資料紀錄】。

經 a200.py 程式，我在【漢字注音】工作表中的每個【漢字】，程式會自動代為查找【台語音標】並填入。但因漢字一字多音，及閩南話之漢字尚分【文讀音】、【白話音】兩種發音，故程式自動查找的【台語音標】往往不對。所以，使用者需要自行校對。

a250.py 程式的功能，在於提供使用者，校對【漢字注音】工作表，發現某【漢字】之【台語音標】有錯時，可依以下程序完成【訂正作業】：

 1. 在【漢字注音】工作表，將有錯之漢字設定為【作用儲存格】；
 2. 在【終端機】執行 a250.py 程式；
 3. 程式執行漢字查尋作業，並條列於 Console 供使用者挑選；
 4. 使用者選擇某【台語音標】，或自行手動輸入【台語音標】；
 5. 程式依【漢字】及【台語音標】，在【標音字庫】工作表找到已有的【資料紀錄】；
 6. 程式依找到的【資料紀錄】，更新【台語音標】欄的舊資料；
 7. 程式依找到的【資料紀錄】，取出【座標】欄的資料，作為：【座標清單】。然後依【座標清單】指向【漢字注音】工作表，各【漢字】儲存格之【座標】，一一為之更新【台語音標】及【漢字標音】。

函數：_update_piau_im_ji_khoo_worksheet() ，負責上述【訂正作業】之【步驟5-7】工作。

更新紀錄：
 -  v0.2.7 2026-02-10: 查得漢字之標音並選用後，因更新【標音字庫】工作表中對應
    【漢字】之【台語音標】及【漢字標音】，導致【作用儲存格】之 Excel Address
    已變更，需將之校正回歸。
 - v0.2.8 2026-02-15: 修正 _process_cell(), _process_jin_kang_piau_im() 函式之使用方式。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# from mod_excel_access import (
#     excel_address_to_row_col,
#     get_active_cell_address,
#     get_line_no_by_row,
#     get_row_by_line_no,
# )
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)

# 載入自訂模組/函式
from mod_標音 import is_han_ji
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
# 資料類別：儲存處理配置
# =========================================================================
class CellProcessor(ExcelCell):
    """
    個人字典查詢專用的儲存格處理器
    繼承自 ExcelCell
    覆蓋以下方法以實現個人字典查詢功能：
    - _process_cell(): 處理單一儲存格
    - _process_jin_kang_piau_im(): 處理人工標音邏輯
    其他方法繼承自父類別 ExcelCell
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

    # def _ca_ji_tian_au_thiam_jin_kang_piau_im(self, active_cell):
    #     """查字典後填入工標音"""
    #     piau_im_huat = self.program.piau_im_huat
    #     piau_im = self.program.piau_im
    #     tai_gi_im_piau = ""
    #
    #     # 依據【作用儲存格】之【漢字】，從【自用字典】查詢【台語音標】
    #     tai_gi_im_piau = self._han_ji_ca_piau_im_kap_cu_tik(active_cell)
    #     if tai_gi_im_piau is None:
    #         return None, None
    #
    #     # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
    #     han_ji_piau_im = tlpa_tng_han_ji_piau_im(
    #         piau_im=piau_im,
    #         piau_im_huat=piau_im_huat,
    #         tai_gi_im_piau=tai_gi_im_piau,
    #     )
    #
    #     # active_cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
    #     # active_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
    #     # active_cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音
    #
    #     return tai_gi_im_piau, han_ji_piau_im

    # def _update_piau_im_ji_khoo_worksheet(
    #     self,
    #     row: int,
    #     col: int,
    #     han_ji: str,
    #     tai_gi_im_piau: str,
    #     han_ji_piau_im: str,
    # ) -> None:
    #     """
    #     訂正作業步驟 5–7：更新【標音字庫】，並同步【漢字注音】工作表。
    #
    #     5. 依【漢字】與作用儲存格【座標】，在【標音字庫】找到既有資料紀錄
    #        （該紀錄仍保留訂正前之【台語音標】）。
    #     6. 將該紀錄之【台語音標】更新為使用者選定／輸入之新音標。
    #     7. 取出該紀錄【座標】欄之【座標清單】，逐一更新【漢字注音】工作表
    #        對應儲存格之【台語音標】與【漢字標音】。
    #     """
    #     # row = cell.row  # 漢字儲存格所在之 Row
    #     # col = cell.column
    #     # han_ji = cell.value
    #     # # 使用者已選定／輸入之新【台語音標】與轉換後之【漢字標音】
    #     # tai_gi_im_piau = cell.offset(-1, 0).value
    #     # han_ji_piau_im = cell.offset(1, 0).value
    #
    #     # ------------------------------------------------------------------
    #     # 步驟 5：依【漢字】與【座標】，在【標音字庫】找到既有資料紀錄
    #     # ------------------------------------------------------------------
    #     row_no = self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
    #         han_ji=han_ji,
    #         coordinate=(row, col),
    #     )
    #     if row_no == -1:
    #         print(f"⚠️  【標音字庫】查無【{han_ji}】座標 ({row}, {col}) 之紀錄，略過字庫更新。")
    #         return
    #
    #     # ------------------------------------------------------------------
    #     # 步驟 6–7：更新【台語音標】，並依【座標清單】同步【漢字注音】工作表
    #     # ------------------------------------------------------------------
    #     self.update_piau_im_worksheet_entry(
    #         coordinate=(row, col),
    #         han_ji=han_ji,
    #         tai_gi_im_piau=tai_gi_im_piau,
    #         han_ji_piau_im=han_ji_piau_im,
    #         piau_im_ji_khoo_dict=self.piau_im_ji_khoo_dict,
    #         row_no=row_no,
    #     )

    # def _update_jin_kang_piau_im_ji_khoo_worksheet(self, cell) -> None:
    #     """
    #     更新【人工標音字庫】工作表
    #     """
    #     row = cell.row
    #     col = cell.column
    #     han_ji = cell.value
    #     tai_gi_im_piau = cell.offset(-1, 0).value

    #     # -------------------------------------------------------------------------
    #     # 確認【漢字】在【人工標音字庫】之【字庫表】，沒有留下舊記錄
    #     # -------------------------------------------------------------------------
    #     row_no = self.jin_kang_piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
    #         han_ji=han_ji,
    #         coordinate=(row, col),
    #     )
    #     # 若是在【標音字庫】中留有舊記錄，需將之移除
    #     if row_no != -1:
    #         # self.jin_kang_piau_im_ji_khoo_dict.remove_coordinate(
    #         #     han_ji=han_ji,
    #         #     coordinate=(row, col),
    #         # )
    #         self.jin_kang_piau_im_ji_khoo_dict.remove_coordinate_by_han_ji_and_coordinate(
    #             han_ji=han_ji,
    #             coordinate=(row, col),
    #         )
    #     # -------------------------------------------------------------------------
    #     # 在【人工標音字庫】之【字庫表】(dict)，新增該【漢字】之記錄
    #     # -------------------------------------------------------------------------
    #     self.jin_kang_piau_im_ji_khoo_dict.add_entry(
    #         han_ji=han_ji,
    #         tai_gi_im_piau=tai_gi_im_piau,
    #         hau_ziann_im_piau='N/A',
    #         coordinate=(row, col),
    #     )
    #     # ----------------------------------------------------------------------
    #     # 將【人工標音字庫】之【字庫表】，寫回 Excel 工作表
    #     # ----------------------------------------------------------------------
    #     self.jin_kang_piau_im_ji_khoo_dict.write_to_excel_sheet(
    #         wb=self.program.wb, sheet_name=self.jin_kang_piau_im_ji_khoo_dict.name
    #     )

    def _process_sheet(self):
        try:
            # --------------------------------------------------------------------------
            # 指定【漢字注音】工作表為【作用工作表】
            # --------------------------------------------------------------------------
            source_sheet_name = self.program.hanji_piau_im_sheet_name
            wb = self.program.wb
            source_sheet = wb.sheets[source_sheet_name]
            source_sheet.activate()

            # ----------------------------------------------------------------------
            # 取得【作用儲存格】
            # ----------------------------------------------------------------------
            # 取得【漢字標音】工作表的【作用儲存格】
            han_ji_cell = self.get_han_ji_cell_with_active_cell()
            # 自【漢字】儲存格取得【位址】（Excel Address）及【座標】 (row, col)
            active_cell_address = han_ji_cell.address.replace("$", "")
            han_ji_row, col = han_ji_cell.row, han_ji_cell.column

            # 確認【作用儲存格】為【漢字】
            han_ji = han_ji_cell.offset(0, 0).value
            if not is_han_ji(han_ji):
                msg = f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
                print(f">> {msg}")
                return EXIT_CODE_SUCCESS

            # 確認【作用儲存格】的【漢字】有【台語音標】及【漢字標音】，否則可能是字典目前無此【漢字】之讀音資料，
            # 故後續之查字典作業應被略過，直接要求使用者輸入【台語音標】或【台羅拼音】。
            jin_kang_piau_im = han_ji_cell.offset(-2, 0).value
            tai_gi_im_piau = han_ji_cell.offset(-1, 0).value
            han_ji_piau_im = han_ji_cell.offset(1, 0).value

            original_tai_gi_im_piau = tai_gi_im_piau
            original_han_ji_piau_im = han_ji_piau_im

            # 確認【作用儲存格】的【台語音標】、【漢字標音】，需已填入資料。
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
            else:
                # ----------------------------------------------------------------------
                # 使用【漢字】查字典，取得【台語音標】
                # ----------------------------------------------------------------------
                han_ji_position = (han_ji_row, col)
                print(f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}")
                print(f"📌 漢字：{han_ji}")
                print(f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}")

                # 查字典後，將查尋所得之【台語音標】，轉換【漢字標音】，最後回傳
                tai_gi_im_piau, han_ji_piau_im = self._ca_ji_tian_au_thiam_jin_kang_piau_im(
                    active_cell=han_ji_cell,
                )
                # 若使用者放棄變更，則結束作業流程
                if not tai_gi_im_piau and not han_ji_piau_im:
                    return EXIT_CODE_SUCCESS

            msg = f"【{han_ji}】[{original_tai_gi_im_piau}] / [{original_han_ji_piau_im}] 變更為： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
            print(f">> 儲存格：{active_cell_address}，{msg}")

            # -------------------------------------------------------------------------
            # 變更【漢字】之讀音：依【作用儲存格】之【人工標音】，變更【漢字標音】工作表
            # 已登錄之【資料紀錄】。使用【人工標音】的漢字讀音，取代【台語音標】欄的資料。
            # -------------------------------------------------------------------------
            self._replace_han_ji_thok_im_by_active_cell(
                row=han_ji_row,
                col=col,
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                han_ji_piau_im=han_ji_piau_im,
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
            han_ji_cell.select()  # 選取【作用儲存格】，以確保游標位置正確

            logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
            return EXIT_CODE_SUCCESS
        except Exception as e:
            # 你可以在這裡加上紀錄或處理，例如:
            logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
            # 再次拋出異常，讓外層函式能捕捉
            raise


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
        # wb.sheets[sheet_name].activate()
        xls_cell._process_sheet()

    except Exception as e:
        logging.error(f"處理錯誤：{e}")
        print(f"❌ 錯誤：{e}")

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
    # (1) 開始執行程式
    # =========================================================================
    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
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
        logging_exc_error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 要求畫面回到【漢字注音】工作表
        # wb.sheets['漢字注音'].activate()
        # 儲存檔案
        wb.save()
        file_path = wb.fullname
        logging_process_step(f"儲存檔案至路徑：{file_path}")

    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束作業
    # =========================================================================
    return EXIT_CODE_SUCCESS


# =========================================================================
# 單元測試程式
# =========================================================================
def test_01():
    """測試 HanJiTian 類別"""
    pass


if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="透過【作用儲格】，查詢漢字之【台語音標】，及生成【漢字標音】",
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
        test_01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            print(f"程式異常終止，返回代碼：{exit_code}")
            sys.exit(exit_code)
