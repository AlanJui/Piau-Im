"""
a260_依字典查得結果填入人工標音.py V0.0.2

在【漢字注音】工作表之【作用儲存格】，可以兩種方式輸入【人工標音】資料：
（1）自【自用字典】查得【台語音標】；（2）直接手動輸入【台語音標】/【台羅拼音】。

修改紀錄：
v0.0.1 2026-2-28: 初始版本，完成基本功能。
v0.0.2 2026-3-21: 修正查字典時，顯示所有讀音的預設值為 True。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from mod_excel_access import (
    excel_address_to_row_col,
    get_active_cell,
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

    def _han_ji_ca_piau_im_kap_cu_tik(self, cell):
        """
        漢字查標音與取得： 依【漢字】從【字典】查得【台語音標】，供使用者挑選：
        （1）字典裡已有的讀音選項；或（2）直接輸入【台語音標】或【台羅拼音】。
        """
        han_ji = cell.value
        tai_gi_im_piau = ""

        if han_ji == "":
            return tai_gi_im_piau

        # (1) 查字典：使用 HanJiTian 類別查詢漢字讀音
        result = self.program.ji_tian.han_ji_ca_piau_im(
            han_ji=han_ji,
            ue_im_lui_piat=self.program.ue_im_lui_piat,
            display_all_piau_im=True,
        )

        # 查無此字
        if not result:
            print(f">> 漢字【{han_ji}】查不到讀音資料！")
            return tai_gi_im_piau

        # (2) 在 console 列出字典中，查詢之漢字有那些讀音選項及其常用程度

        # 顯示所有讀音選項
        piau_im_options = self.display_all_piau_im_for_a_han_ji(han_ji, result)

        # (3) 供使用者輸入選擇
        user_input = (
            input("\n請輸入選擇編號 (直接按 Enter 跳過): ").strip().lstrip("\ufeff")
        )

        if not user_input:
            print(">> 放棄變更！")
            return None

        try:
            # 取得使用者之輸入，並【解析】其輸入是要：（1）引用字典的查找結果；
            # （2）直接輸入【台語音標】或【台羅拼音】
            if user_input.isdigit():
                choice = int(user_input)
                case = 1
            else:
                case = 2

            if case == 1:
                # （1）引用字典查找結果
                if 1 <= choice <= len(piau_im_options):
                    # 顯示使用者輸入之讀音選項
                    print(f"【{han_ji}】讀音，選用：第 {choice} 個選項。")

                    # 依據輸入之【數值】，自讀音選項清單(piau_im_options)，取得對映之【台語音標】及【漢字標音】
                    selected_im_piau, selected_han_ji_piau_im = piau_im_options[
                        choice - 1
                    ]

                    # return [selected_im_piau, selected_han_ji_piau_im]
                    return selected_im_piau
                else:
                    print(f">> 輸入錯誤：{choice} 超出範圍！")
                    return None
            elif case == 2:
                # （2）直接輸入【台語音標】或【台羅拼音】
                raw_im_piau = user_input.lower()
                import re

                if re.match(
                    r"^[a-zâîûêôáéíóúàèìòùāēīōūǎěǐǒǔ]+[1-8]?$", raw_im_piau, re.I
                ):
                    print(f"【{han_ji}】讀音，採用直接輸入：【{raw_im_piau}】")
                    return raw_im_piau
                else:
                    print(f">> 輸入格式有誤：【{raw_im_piau}】不是有效的羅馬拼音格式！")
                    return None
        except ValueError:
            print(f">> 使用者輸入格式有誤：{user_input}")
            return None

        return tai_gi_im_piau

    def _za_ji_tain_au_thiam_jin_kang_piau_im(self, active_cell):
        """查字典後填入工標音"""
        piau_im_huat = self.program.piau_im_huat
        piau_im = self.program.piau_im
        tai_gi_im_piau = ""

        # 依據【作用儲存格】之【漢字】，從【自用字典】查詢【台語音標】
        tai_gi_im_piau = self._han_ji_ca_piau_im_kap_cu_tik(active_cell)
        if tai_gi_im_piau is None:
            return

        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau,
        )

        active_cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
        active_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        active_cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音

        return tai_gi_im_piau, han_ji_piau_im

    def _manual_input_thok_im(self, cell):
        """
        手動輸入【漢字】之【台語音標】讀音

        returns:
            tai_gi_im_piau: 使用者輸入的台語音標
            han_ji_piau_im: 由台語音標轉換而成的漢字標音

            cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
            cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
            cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音
        """
        # 取得【漢字】儲存格內容
        han_ji = cell.value

        tai_gi_im_piau = self.get_user_input_piau_im(han_ji=han_ji)
        if not tai_gi_im_piau:
            return
        cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標

        # 將使用者輸入之【台語音標】轉換為【漢字標音】
        han_ji_piau_im = self.convert_tai_gi_im_piau_to_han_ji_piau_im(
            tai_gi_im_piau=tai_gi_im_piau,
        )
        if not han_ji_piau_im:
            print(">> 無法將輸入之【台語音標】轉換為【漢字標音】！")
            return tai_gi_im_piau, None

        cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
        cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音

        return tai_gi_im_piau, han_ji_piau_im

    def _update_khuat_ji_piau_worksheet(self, cell) -> None:
        """
        更新【缺字表】工作表
        """
        row = cell.row
        col = cell.column
        han_ji = cell.value
        tai_gi_im_piau = cell.offset(-1, 0).value
        han_ji_piau_im = cell.offset(1, 0).value

        # 在【缺字表】工作表查找此【漢字】之 Excel 的 Row No
        row_no = self.khuat_ji_piau_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
            han_ji=han_ji,
            coordinate=(row, col),
        )
        if row_no != -1:
            # 找到【漢字】所在之 Row No 後，依據【座標】欄儲存格之【座標清單】，逐一更新指向
            # 【漢字注音】工作表之【漢字】的【台語音標】及【漢字標音】。
            # 之【台語音標】及【漢字標音】。
            self.update_piau_im_worksheet_entry(
                coordinate=(row, col),
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                han_ji_piau_im=han_ji_piau_im,
                piau_im_ji_khoo_dict=self.khuat_ji_piau_ji_khoo_dict,
                row_no=row_no,
            )
            # 因【標音字庫】依【漢字】之【座標】紀錄，更新【漢字注音】工作表中對映之【台語音標】及【漢字標音】；導致
            # 【作用儲存格】之 Excel Address 已變更，需將之校正回歸。
            cell.select()  # 選取【作用儲存格】，以確保游標位置正確

    def _update_jin_kang_piau_im_ji_khoo_worksheet(self, cell) -> None:
        """
        更新【人工標音字庫】工作表
        """
        row = cell.row
        col = cell.column
        han_ji = cell.value
        tai_gi_im_piau = cell.offset(-1, 0).value

        # 是否存在【標音字庫】
        row_no = self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
            han_ji=han_ji,
            coordinate=(row, col),
        )
        # 若是存在【標音字庫】之中，需移除
        if row_no != -1:
            self.piau_im_ji_khoo_dict.remove_entry(
                han_ji=han_ji,
                coordinates=(row, col),
            )

        # 將【人工標音】記錄到【人工標音字庫】工作表
        self.jin_kang_piau_im_ji_khoo_dict.add_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau='N/A',
            coordinates=(row, col),
        )

# =============================================================================
# 作業主流程
# =============================================================================
def process(wb, args) -> int:
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】

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
        # 初始化 process config
        # --------------------------------------------------------------------------
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        # --------------------------------------------------------------------------
        # 處理作業開始
        # --------------------------------------------------------------------------
        source_sheet_name = "漢字注音"
        jin_kang_piau_im_sheet_name = "人工標音字庫"
        piau_im_ji_khoo_sheet_name = "標音字庫"

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
        jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = (
            get_row_by_line_no(current_line_no)
        )

        # 確認【作用儲存格】為【漢字】
        han_ji = source_sheet.range((han_ji_row, col)).value
        if not is_han_ji(han_ji):
            msg=f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_SUCCESS

        # 確認【作用儲存格】的【漢字】有【台語音標】及【漢字標音】，否則可能是字典目前無此【漢字】之讀音資料，
        # 故後續之查字典作業應被略過，直接要求使用者輸入【台語音標】或【台羅拼音】。
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value

        if not tai_gi_im_piau or not han_ji_piau_im:
            # ----------------------------------------------------------------------
            # 直接手動輸入人工標音，若是【作用儲存格】之【漢字】，可能字典尚未登錄此漢字之讀音資料
            # ----------------------------------------------------------------------
            msg = f"作用儲存格 {active_cell_address} 的漢字【{han_ji}】缺乏【台語音標】或【漢字標音】，可能是字典無此漢字之讀音資料，將略過查字典作業，直接要求使用者輸入【台語音標】或【台羅拼音】。"
            print(f">> {msg}")
            # 取得使用者輸入之【台語音標】或【台羅拼音】
            tai_gi_im_piau = xls_cell.get_user_input_piau_im(han_ji=han_ji)
            # 依據使用者輸入之【台語音標】轉換為【漢字標音】
            han_ji_piau_im = xls_cell._convert_tai_gi_im_piau_to_han_ji_piau_im(
                tai_gi_im_piau=tai_gi_im_piau,
            )

            source_sheet.range((tai_gi_im_piau_row, col)).value = tai_gi_im_piau
            source_sheet.range((han_ji_piau_im_row, col)).value = han_ji_piau_im
            source_sheet.range((jin_kang_piau_im_row, col)).value = jin_kang_piau_im
        else:
            # ----------------------------------------------------------------------
            # 查字典後填人工標音
            # ----------------------------------------------------------------------
            han_ji_position = (han_ji_row, col)
            print(
                f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}"
            )
            print(f"📌 漢字：{han_ji}")
            print(
                f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}"
            )
            xls_cell._za_ji_tain_au_thiam_jin_kang_piau_im(active_cell=active_cell)

        # 透過【作用儲存格】取出處理後的【人工標音】、【台語音標】、【漢字標音】
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value

        msg = f"{han_ji}： [{jin_kang_piau_im}] / [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
        print(f">> 儲存格：{active_cell_address}，變更結果為：{msg}")

        # 將【台語音標】和【漢字標音】寫入【漢字注音】工作表之【作用儲存格】
        if not jin_kang_piau_im:    # 若【人工標音】儲存格未填入標音
            return EXIT_CODE_SUCCESS

        # -------------------------------------------------------------------------
        # 自【標音字庫】之【字庫表】(dict)，移除該【漢字】之記錄
        # -------------------------------------------------------------------------
        # 調整 row 指向【漢字】儲存格所在座標列
        row = han_ji_row

        # 是否存在【標音字庫】
        row_no = xls_cell.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
            han_ji=han_ji,
            coordinate=(row, col),
        )
        # 若是存在【標音字庫】之中，需移除
        if row_no != -1:
            xls_cell.piau_im_ji_khoo_dict.remove_entry(
                han_ji=han_ji,
                coordinates=(row, col),
            )
        # -------------------------------------------------------------------------
        # 在【人工標音字庫】之【字庫表】(dict)，新增該【漢字】之記錄
        # -------------------------------------------------------------------------
        # 確認【漢字】在【人工標音字庫】之【字庫表】，沒有留下舊記錄
        row_no = xls_cell.jin_kang_piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
            han_ji=han_ji,
            coordinate=(row, col),
        )
        if row_no != -1:
            xls_cell.jin_kang_piau_im_ji_khoo_dict.remove_coordinate(
                han_ji=han_ji,
                coordinates=(row, col),
            )
        # 在【人工標音字庫】之【字庫表】，添加標音記錄
        xls_cell.jin_kang_piau_im_ji_khoo_dict.add_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau='N/A',
            coordinates=(row, col),
        )
        # ----------------------------------------------------------------------
        # 將【標音字庫】之【字庫表】，寫回 Excel 工作表
        # ----------------------------------------------------------------------
        xls_cell.piau_im_ji_khoo_dict.write_to_excel_sheet(
            wb=wb, sheet_name=piau_im_ji_khoo_sheet_name
        )
        # ----------------------------------------------------------------------
        # 將【人工標音字庫】之【字庫表】，寫回 Excel 工作表
        # ----------------------------------------------------------------------
        xls_cell.jin_kang_piau_im_ji_khoo_dict.write_to_excel_sheet(
            wb=wb, sheet_name=jin_kang_piau_im_sheet_name
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

        logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
        # 再次拋出異常，讓外層函式能捕捉
        raise


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
        msg = f"程式異常終止：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
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


def ut01():
    # 取得【作用中活頁簿】
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
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
    args = parser.parse_args()

    if args.test:
        # 執行測試
        ut01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code == EXIT_CODE_SUCCESS:
            print("程式正常完成！")
        else:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
