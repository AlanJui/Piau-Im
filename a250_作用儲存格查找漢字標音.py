"""
a250_作用儲存格查找漢字標音.py V0.2.4

透過在【漢字注音】工作表點選【作用儲存格】，便能查詢儲存格所對映之漢字的
【台語音標】，及生成使用者慣用之【漢字標音】。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import re
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)

# 載入自訂模組
from mod_標音 import is_punctuation
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
# 輔助函數
# =========================================================================
def _get_active_cell_from_sheet(sheet, xls_cell: ExcelCell):
    """處理整個工作表"""
    program = xls_cell.program

    # 自【作用儲存格】取得【Excel 儲存格座標】(列,欄) 座標
    active_cell = sheet.api.Application.ActiveCell
    if active_cell:
        # 顯示【作用儲存格】位置
        active_row = active_cell.Row
        active_col = active_cell.Column
        active_col_name = xw.utils.col_name(active_col)
        print(
            f"作用儲存格：{active_col_name}{active_row}（{active_cell.Row}, {active_cell.Column}）"
        )

        # 調整 row 值至【漢字】列（每 4 列為一組，漢字在第 3 列：5, 9, 13, ... ）
        line_start_row = 3  # 第一行【標音儲存格】所在 Excel 列號: 3
        line_no = (active_row - line_start_row + 1) // program.ROWS_PER_LINE
        row = (line_no * program.ROWS_PER_LINE) + xls_cell.program.han_ji_row_offset - 1
        col = active_cell.Column
        cell = sheet.range((row, col))
        # cell.select()

        # 處理儲存格
        xls_cell._process_cell(cell, row, col)


# =========================================================================
# 主要處理函數
# =========================================================================
class CellProcessor(ExcelCell):
    """
    本程式專用的儲存格處理器
    繼承自 ExcelCell 的類別
    覆蓋以下方法以實現萌典查詢功能：
    - _process_han_ji(): 使用【個人字典】查詢漢字讀音
    - _process_cell(): 處理單一儲存格
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
    # 覆蓋父類別的方法
    # =================================================================

    def _process_jin_kang_piau_im(
        self, han_ji: str, jin_kang_piau_im: str, cell, row: int, col: int
    ):
        """處理人工標音內容"""
        # 預設未能依【人工標音】欄，找到對應的【台語音標】和【漢字標音】
        # org_tai_gi_im_piau = cell.offset(-1, 0).value
        han_ji = cell.value

        # 取得【漢字】儲存格之【座標】位址（row, col）
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

    def check_coordinate_exists(
        self,
        row: int,
        col: int,
        coord_list,
    ) -> bool:
        """
        檢查指定座標是否存在於座標列表中。
        Args:
            row (int): 要檢查的列號
            col (int): 要檢查的欄號
            coord_list: 座標列表（可以是 list 或 str）
        Returns:
            bool: 如果座標存在則返回 True，否則返回 False
        """
        # 如果是字串格式，先解析成 list
        if isinstance(coord_list, str):
            # coord_str = '(61, 13); (69, 8); (89, 13); (125, 11); (125, 16)'
            coordinate_tuples = re.findall(r"\((\d+)\s*,\s*(\d+)\)", coord_list)
            coordinate_list = [(int(r), int(c)) for r, c in coordinate_tuples]
        else:
            # 已經是 list 格式：[(61, 13), (69, 8), (89, 13), (125, 11), (125, 16)]
            coordinate_list = coord_list

        # 判斷是否存在
        coordinate = (row, col)
        exists = coordinate in coordinate_list

        # print(f"座標 {coordinate} 存在: {exists}")  # True
        # print(f"所有座標: {coordinate_list}")
        return exists

    def _process_han_ji(
        self,
        han_ji: str,
        cell,
        row: int,
        col: int,
    ) -> str:
        """
        處理漢字 - 使用【個人字典】查詢讀音
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
            return "【空白】"

        # 使用 HanJiTian 查詢漢字讀音
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
            return f"【{han_ji}】查無此字！"

        # 有多個讀音
        print(
            # f"漢字儲存格：{xw.utils.col_name(col)}{row}（{row}, {col}）：【{han_ji}】有 {len(result)} 個讀音：{result}"
            f"【{han_ji}】有 {len(result)} 個讀音：{result}"
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
            print("-" * 80)
            print(f"{idx + 1}. {msg}")

        # 讓使用者選擇讀音
        user_input = input(
            "\n請選擇讀音編號（直接按 Enter 略過，輸入編號後按 Enter 填入）："
        ).strip()

        if user_input == "":
            # 只瀏覽，不填入
            print(f"【{han_ji}】略過填注標音！")
            return None

        try:
            choice = int(user_input)
            if 1 <= choice <= len(result):
                # 填入選擇的讀音
                tai_gi_im_piau, han_ji_piau_im = piau_im_options[choice - 1]
                # cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音儲存格
                # cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標儲存格
                # cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音儲存格

                # 在【標音字庫】查找是否此【漢字】，已有指向【漢字注音】工作表之【座標】紀錄
                row_no = self.piau_im_ji_khoo_dict.get_row_by_han_ji_and_coordinate(
                    han_ji=han_ji, coordinate=(row, col)
                )
                if row_no != -1:
                    # 已有紀錄，無需新增
                    _, entry = self.piau_im_ji_khoo_dict.get_entry_by_row_no(row_no)
                    exist = self.check_coordinate_exists(
                        row=row,
                        col=col,
                        coord_list=entry["coordinates"],
                    )
                    if exist:
                        print(
                            f"\n【{han_ji}】在【標音字庫】工作表已有 {len(entry['coordinates'])} 筆紀錄，指向【漢字注音】工作表！"
                        )
                        print(f"({row}, {col}) ==> {entry['coordinates']}\n")

                        # 依據【座標】儲存格取得之【座標清單】，逐一變更各組座標所指向【漢字注音】工作表
                        # 之漢字所對映之讀音為【台語音標】、【漢字標音】
                        sheet = self.program.wb.sheets["漢字注音"]
                        i = 1
                        for coordinate in entry["coordinates"]:
                            row, col = coordinate
                            target_cell = sheet.range((row, col))
                            target_cell.select()
                            target_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
                            target_cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音
                            excel_addr = xw.utils.col_name(col) + str(row)
                            print("-" * 80)
                            print(
                                f"已更新 ==> 第 {i} 個： {excel_addr}  ({row}, {col}) 【{han_ji}】：【{tai_gi_im_piau}] /【{han_ji_piau_im}】"
                            )
                            i += 1
                        print(f"【{han_ji}】已填入第 {choice} 個讀音！")

                        # 更新【標音字庫】工作表之【台語音標】
                        self.piau_im_ji_khoo_dict.update_whole_entry(
                            row_no=row_no,
                            tai_gi_im_piau=tai_gi_im_piau,
                            hau_ziann_im_piau="N/A",
                            coordinates=entry["coordinates"],
                        )
                        # 儲存回標音字庫工作表
                        self.save_all_piau_im_ji_khoo_dicts()
                        return None
            else:
                print(f"輸入錯誤：{choice}（超出範圍）")
                return None
        except ValueError:
            print(f"使用者予無用之輸入：{user_input}")
            return None

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
        # self._reset_cell_style(cell)

        # 取得【漢字】儲存格內容
        cell_value = cell.value

        # 檢查是否有【人工標音】
        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
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

        # 依據【漢字】儲存格內容進行處理
        if cell_value == "φ":
            print("【文字終結】")
            return 1  # 文章終結符號
        elif cell_value == "\n":
            print("【換行】")
            return 2  # 【換行】
        elif cell_value is None or str(cell_value).strip() == "":
            print("【空白】")
            return 3  # 空白或標點符號
        elif is_punctuation(cell_value):
            msg = self._process_non_han_ji(cell_value)
            print(msg)
            return 3  # 空白或標點符號
        else:
            # msg = self._process_han_ji(cell_value, cell, row, col)
            # print(msg)
            self._process_han_ji(cell_value, cell, row, col)
            return 0  # 漢字

    def _process_sheet(self, sheet):
        """處理整個工作表"""
        # 取得【作用儲存格】
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


def process(wb, args) -> int:
    """
    查詢漢字讀音並標注

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    # logging_process_step("<=========== 作業開始！==========>")

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

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    print("=" * 80)
    print("無限循環模式：請在 Excel 中選擇任一儲存格後按 Enter 查詢")
    print("按 Ctrl+C 終止程式")
    print("=" * 80)

    # 無限循環
    while True:
        try:
            # 等待使用者按 Enter
            input("\n請在 Excel 選擇【作用儲存格】後按 Enter 繼續（Ctrl+C 終止）...")

            # 確保工作表為作用中
            wb.sheets[sheet_name].activate()

            xls_cell._process_sheet(sheet=sheet)
            print("=" * 80)
            print("\n")

        except KeyboardInterrupt:
            print("\n\n使用者中斷程式（Ctrl+C）")
            print("=" * 70)
            break  # 中斷循環

        except Exception as e:
            logging.error(f"處理錯誤：{e}")
            print(f"❌ 錯誤：{e}")
            # 發生錯誤時繼續循環，不中斷程式
            continue
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
