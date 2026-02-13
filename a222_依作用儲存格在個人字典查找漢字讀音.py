"""
a222_依作用儲存格在個人字典查找漢字讀音.py V0.2.9
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

    def ca_tai_gi_im_piau(self, han_ji: str, cell) -> list[str] | None:
        """查字典"""
        if han_ji == "":
            return None

        # (1) 查字典：使用 HanJiTian 類別查詢漢字讀音
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
                coordinates=(cell.row, cell.col),
            )
            return None

        # (2) 在 console 列出字典中，查詢之漢字有那些讀音選項及其常用程度

        # 顯示所有讀音選項
        piau_im_options = self.display_all_piau_im_for_a_han_ji(han_ji, result)

        # (3) 供使用者輸入選擇
        user_input = input("\n請輸入選擇編號 (直接按 Enter 跳過): ").strip()

        if not user_input:
            print(">> 放棄變更！")
            return None

        try:
            choice = int(user_input)
            if 1 <= choice <= len(piau_im_options):
                # 顯示使用者輸入之讀音選項
                print(f"【{han_ji}】讀音，選用：第 {choice} 個選項。")

                # 依據輸入之【數值】，自讀音選項清單(piau_im_options)，取得對映之【台語音標】及【漢字標音】
                selected_im_piau, selected_han_ji_piau_im = piau_im_options[choice - 1]

                return [selected_im_piau, selected_han_ji_piau_im]
            else:
                print(f">> 輸入錯誤：{choice} 超出範圍！")
                return None
        except ValueError:
            print(f">> 使用者輸入格式有誤：{user_input}")
            return None

    # =================================================================
    # 覆蓋父類別的方法
    # =================================================================
    def _process_jin_kang_piau_im(self, cell):
        """處理人工標音內容"""
        # 預設未能依【人工標音】欄，找到對應的【台語音標】和【漢字標音】
        # org_tai_gi_im_piau = cell.offset(-1, 0).value
        han_ji = cell.value
        jin_kang_piau_im = cell.offset(-2, 0).value
        tai_gi_im_piau = cell.offset(-1, 0).value
        han_ji_piau_im = cell.offset(1, 0).value

        # 取得【漢字】儲存格之【座標】位址（row, col）
        row = cell.row  # 取得【漢字】儲存格的列號
        col = cell.column  # 取得【漢字】儲存格的欄號
        han_ji_row, han_ji_col = self.get_han_ji_coordinate_by_row_and_col(
            row=row, col=col
        )

        # 判斷【人工標音】是要【引用既有標音】還是【手動輸入標音】
        if jin_kang_piau_im == "=":  # 引用既有的人工標音
            tai_gi_im_piau, han_ji_piau_im = self.in_iong_jin_kang_piau_im_ji_khoo(
                han_ji=han_ji,
                jin_kang_piau_im=jin_kang_piau_im,
                cell=cell,
                row=han_ji_row,
                col=han_ji_col,
            )
        elif jin_kang_piau_im == "#":  # 清除人工標音，回復自動標音（使用【標音字庫】）
            # 自【標音字庫】工作表，取得對應的【台語音標】和【漢字標音】
            tai_gi_im_piau, han_ji_piau_im = self.in_iong_piau_im_ji_khoo(
                han_ji=han_ji,
                jin_kang_piau_im=jin_kang_piau_im,
                cell=cell,
                row=han_ji_row,
                col=han_ji_col,
            )
        else:  # 自【人工標音】儲存格，解析【人工標音】輸入之【台語音標】或【台羅拼音】
            tai_gi_im_piau, han_ji_piau_im = self._cu_jin_kang_piau_im(
                jin_kang_piau_im=str(jin_kang_piau_im),
                piau_im=self.program.piau_im,
                piau_im_huat=self.program.piau_im_huat,
            )
            if tai_gi_im_piau != "" and han_ji_piau_im != "":
                # 自【標音字庫】工作表，移除【漢字】及指向【漢字注音】工作表之【座標】
                self.piau_im_ji_khoo_dict.remove_coordinate_by_han_ji_and_coordinate(
                    han_ji=han_ji, coordinate=(han_ji_row, han_ji_col)
                )
                # 在【人工標音字庫】新增一筆資料，記錄：【漢字】、【台語音標】及指向【漢字注音】之【座標】
                self.jin_kang_piau_im_ji_khoo_dict.add_or_update_entry(
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    hau_ziann_im_piau="N/A",
                    coordinates=(han_ji_row, han_ji_col),
                )
                # ---------------------------------------------------------------------------------
                # 顯示處理訊息
                # ---------------------------------------------------------------------------------
                coordinate_str = None
                # excel_addr = convert_row_col_to_excel_address(row, col)
                # source_msg = f"【漢字注音】工作表 {excel_addr}（{row} ,{col}）==》漢字：【{han_ji}】，人工標音：【{jin_kang_piau_im}】"
                source_msg = f"==》漢字：【{han_ji}】，人工標音：【{jin_kang_piau_im}】"
                print(f"{source_msg} ...")

                # 顯示【人工標音字庫】工作表新增之紀錄
                row_no_jin_kang_piau_im = (
                    self.jin_kang_piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
                        han_ji=han_ji, coordinate=(row, col)
                    )
                )
                if row_no_jin_kang_piau_im:
                    result = self.jin_kang_piau_im_ji_khoo_dict.get_entry_by_row_no(
                        row_no=row_no_jin_kang_piau_im
                    )
                    if result:
                        _, entry = result
                        tai_gi_im_piau = entry.get("tai_gi_im_piau", "")
                        coordinate_list = entry.get("coordinates", [])
                        # 使用 join 轉換（推薦，格式化後的字串）
                        coordinate_str = (
                            "; ".join([f"({r}, {c})" for r, c in coordinate_list])
                            if coordinate_list
                            else "無"
                        )
                    else:
                        coordinate_str = "無"
                else:
                    coordinate_str = "無"
                target_msg = f"在【人工標音字庫】工作表 {row_no_jin_kang_piau_im}A（{row_no_jin_kang_piau_im}, 1）新增一筆紀錄 ==> 漢字：【{han_ji}】，台語音標：【{tai_gi_im_piau}】，座標：【{coordinate_str}】"
                print(f"{target_msg}")

                # 顯示【標音字庫】工作表移除的紀錄
                row_no_piau_im_ji_khoo = (
                    self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
                        han_ji=han_ji, coordinate=(row, col)
                    )
                )
                if row_no_piau_im_ji_khoo:
                    result = self.piau_im_ji_khoo_dict.get_entry_by_row_no(
                        row_no=row_no_piau_im_ji_khoo
                    )
                    if result:
                        _, entry = result
                        coordinate_list = entry.get("coordinates", [])
                        # 使用 join 轉換（推薦，格式化後的字串）
                        coordinate_str = (
                            "; ".join([f"({r}, {c})" for r, c in coordinate_list])
                            if coordinate_list
                            else "無"
                        )
                    else:
                        coordinate_str = "無"
                else:
                    coordinate_str = "無"
                if row_no_piau_im_ji_khoo == -1:
                    target_msg2 = f"原【標音字庫】工作表無漢字：【{han_ji}】之紀錄。"
                else:
                    target_msg2 = f"原【標音字庫】工作表 {row_no_piau_im_ji_khoo}A（{row_no_piau_im_ji_khoo}, 1）移除其【座標】紀錄 ==> 漢字：【{han_ji}】，座標：【{coordinate_str}】"
                print(f"{target_msg2}")

        # 將結果儲存回標音字庫工作表
        self.save_all_piau_im_ji_khoo_dicts()

    def _process_bo_thok_im(self, cell):
        """處理【無標音漢字】的【台語音標】及【漢字標音】儲存格內容"""
        row = cell.row  # 取得【漢字】儲存格的列號
        col = cell.column  # 取得【漢字】儲存格的欄號
        # 取得【漢字】儲存格內容
        han_ji = cell.value

        # 查字典：使用 HanJiTian 類別查詢漢字讀音
        result = self.ca_tai_gi_im_piau(han_ji=han_ji, cell=cell)
        if not result:
            return
        tai_gi_im_piau, han_ji_piau_im = result

        # 更新【台語音標】及【漢字標音】儲存格內容
        cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音

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

    def _process_han_ji(self, cell):
        """
        使用【個人字典】查詢讀音
        (1) 查字典
        (2) 列出選項 (音標、常用度)
        (3) 使用者輸入選擇

        Args:
            cell: Excel 儲存格物件

        Returns:
            (message, success): 處理訊息和是否成功
        """

        row = cell.row  # 取得【漢字】儲存格的列號
        col = cell.column  # 取得【漢字】儲存格的欄號
        # 取得【漢字】儲存格內容
        han_ji = cell.value

        # 查字典：使用 HanJiTian 類別查詢漢字讀音
        result = self.ca_tai_gi_im_piau(han_ji=han_ji, cell=cell)
        if not result:
            return
        tai_gi_im_piau, han_ji_piau_im = result

        # 更新【台語音標】及【漢字標音】儲存格內容
        cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音

        # 在【標音字庫】工作表查找此【漢字】之 Excel 的 Row No
        row_no = self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
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
                piau_im_ji_khoo_dict=self.piau_im_ji_khoo_dict,
                row_no=row_no,
            )
            # 因【標音字庫】依【漢字】之【座標】紀錄，更新【漢字注音】工作表中對映之【台語音標】及【漢字標音】；導致
            # 【作用儲存格】之 Excel Address 已變更，需將之校正回歸。
            cell.select()  # 選取【作用儲存格】，以確保游標位置正確

    def _process_cell(
        self,
        cell,
        row: int,
        col: int,
    ) -> bool:
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

        cell_value = cell.value

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
            msg = self._process_non_han_ji(cell_value)
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

        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        tai_gi_im_piau = cell.offset(-1, 0).value  # 台語音標
        han_ji_piau_im = cell.offset(1, 1).value  # 漢字標音

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
            self._process_jin_kang_piau_im(
                han_ji=cell_value,
                jin_kang_piau_im=jin_kang_piau_im,
                cell=cell,
                row=row,
                col=col,
            )
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
