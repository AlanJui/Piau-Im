"""
a100_作業中活頁檔填入漢字.py V0.2.9
功能：將漢字純文字檔中的漢字，填入 Excel 活頁簿中的【漢字注音】工作表，並自動查找台語音標與漢字標音。
更新紀錄：
 v0.2.7 2026-02-08：
  - 改善操作介面：在填入漢字的過程中，顯示正在處理的段落文字，讓使用者能更清楚目前進度。
  - 改善 total_lines 的計算方式：改為從 Excel 工作表中讀取實際的資料行數，而非使用固定值，提升彈性與適應性。
 v0.2.8 2026-02-15:
  - 與【標音】相關的三張工作表（標音字庫、人工標音字庫、缺字表），在執行 process() 時，務必建立新表
  （刪除舊表、建立新表），避免舊文章之標音資料與之相混。
 v0.2.0.9 2026-02-17:
  - 修正 _process_cell() 呼叫方式，改為傳入 active_cell 物件；
    原 row, col 值之取得，透過 active_cell 物件即可。
"""

# =========================================================================
# 程式功能摘要
# =========================================================================
# 用途：將漢字填入對應的儲存格
# 詳述：待加註讀音的漢字文置於 V3 儲存格。本程式將漢字逐字填入對應的儲存格：
# 【第一列】D5, E5, F5,... ,R5；
# 【第二列】D9, E9, F9,... ,R9；
# 【第三列】D13, E13, F13,... ,R13；
# 每個漢字佔一格，每格最多容納一個漢字。
# 漢字上方的儲存格（如：D4）為：【台語音標】欄，由【羅馬拼音字母】組成拼音。
# 漢字下方的儲存格（如：D6）為：【台語注音符號】欄，由【台語方音符號】組成注音。
# 漢字上上方的儲存格（如：D3）為：【人工標音】欄，可以只輸入【台語音標】；或
# 【台語音標】和【台語注音符號】皆輸入。

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import argparse
import logging
import os
import re
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_excel_access import (
    calculate_total_lines,
    clear_han_ji_kap_piau_im,
    reset_cells_format_in_sheet,
)
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)
from mod_帶調符音標 import read_text_with_han_ji
from mod_標音 import is_han_ji
from mod_程式 import ExcelCell, Program

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
DB_KONG_UN = os.getenv("DB_KONG_UN", "Kong_Un.db")

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
# 主要處理函數
# =========================================================================
def fill_in_han_ji(
    wb, text_with_han_ji: list, sheet_name: str = "漢字注音", start_row: int = 5
):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range("A1").select()

    # ------------------------------------------------------------------------------
    # 填入【漢字】
    # ------------------------------------------------------------------------------
    row_han_ji = start_row  # 漢字位置
    start_col = 4  # 從D欄開始
    max_col = 18  # 最大可填入的欄位（R欄）

    col = start_col

    print(
        f"正在填入漢字到工作表 {sheet_name}，起始列 {start_row}，起始欄 {start_col}，最大欄 {max_col}"
    )
    text = ""

    for han_ji_ku in text_with_han_ji:
        # 1. 取得該行文字
        line_text = "".join(han_ji_ku)
        # 2. 顯示即將處理的文字
        # print(f"處理段落：{line_text}")
        print(f"==> {line_text}")

        for han_ji in han_ji_ku:
            if col > max_col:
                # 超過欄位，換到下一組行
                row_han_ji += 4
                col = start_col

            text += han_ji
            sheet.cells(row_han_ji, col).value = han_ji
            sheet.cells(row_han_ji, col).select()  # 選取，畫面滾動
            col += 1  # 填入後右移一欄
            # 以下程式碼有假設：每組漢字之結尾，必有標點符號

        # 段落終結處：換下一段落
        if col > max_col:
            # 超過欄位，換到下一組行
            row_han_ji += 4
            col = start_col
        sheet.cells(row_han_ji, col).value = "=CHAR(10)"
        text += "\n"

        row_han_ji += 4
        col = start_col

    # 填入文章終止符號：φ
    # sheet["V3"].value = text
    sheet.cells(row_han_ji, col).value = "φ"
    print(f"已將文章之漢字純文字檔讀入，並填進【{sheet_name}】工作表！")

    return text_with_han_ji


def _fill_han_ji_into_sheet(
    wb,
    program: Program,
    text_file_name: str = "_tmp_p1_han_ji.txt",
    sheet_name: str = "漢字注音",
    target: str = "V3",
) -> int:
    """填入【漢字】到指定工作表"""
    # 讀取漢字檔，並填入 Excel
    text_with_han_ji = read_text_with_han_ji(filename=text_file_name)
    text_with_han_ji = fill_in_han_ji(
        wb, text_with_han_ji, sheet_name=sheet_name, start_row=program.han_ji_start_row
    )

    # 建漢字清單：將 text_with_han_ji 整編為【漢字清單】
    han_ji_list = []
    for han_ji_ku in text_with_han_ji:
        for han_ji in han_ji_ku:
            han_ji_list.append(han_ji)
        # 段落終結處：換下一段落
        han_ji_list.append("\n")
    # 將漢字檔已讀取之內容，填入【漢字注音】工作表之【V3】儲存格
    wb.sheets[sheet_name].range(target).value = "".join(han_ji_list)

    # 將文件標題提取並寫入 env 表 TITLE 名稱格
    extract_and_set_title(wb, text_file_name)


def extract_and_set_title(wb, file_path):
    """從漢字純文字檔中提取標題，並寫入 env 表 TITLE 名稱格"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            first_line = f.readline().strip()
            match = re.search(r"《(.*?)》", first_line)
            if match:
                title = match.group(1)
                wb.names["TITLE"].refers_to_range.value = title
                logging.info(f"✅ 已將文件標題《{title}》寫入 env 表 TITLE 名稱格。")
            else:
                logging.info("❕ 無《標題》可提取，未更新 TITLE。")
    except Exception as e:
        logging_exc_error("無法讀取或更新 TITLE 名稱。", error=e)


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
        # tai_gi_im_piau = cell.offset(-1, 0).value  # 台語音標
        # han_ji_piau_im = cell.offset(1, 1).value  # 漢字標音

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
            self._process_non_han_ji(cell)
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
        # if (
        #     not tai_gi_im_piau
        #     or str(tai_gi_im_piau).strip() == ""
        #     and not han_ji_piau_im
        # ):
        #     self._process_bo_thok_im(cell)
        #     return 0  # 漢字

        # 檢查是否為【人工標音漢字】
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            self._show_msg(row, col, cell_value)
            self._process_jin_kang_piau_im(cell=cell)
            return 0  # 漢字

        # 處理【自動標音漢字】
        self._process_han_ji(cell)
        return 0  # 漢字

    def _process_sheet(self, sheet, show_cell_address: bool = False):
        """處理整個工作表"""
        # 初始化變數
        wb = sheet.book
        config = self.program
        total_lines = config.TOTAL_LINES
        line_start_row = config.line_start_row
        rows_per_line = config.ROWS_PER_LINE
        # start_row = line_start_row + 2  # 調整為實際起始列
        # end_row = start_row + (config.TOTAL_LINES * config.ROWS_PER_LINE)
        start_col = config.start_col
        end_col = config.end_col
        han_ji_row_offset = config.han_ji_row_offset

        # 處理所有的儲存格
        active_cell = sheet.range(f"{xw.utils.col_name(start_col)}{line_start_row}")
        active_cell.select()

        # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
        is_eof = False
        # total_lines = program.TOTAL_LINES
        # 計算【漢字注音】工作表的【漢字注音行】總行數
        total_lines = calculate_total_lines(sheet)
        try:
            wb.names["每頁總列數"].refers_to_range.value = total_lines
        except Exception:
            pass  # 若無此名稱定義，則忽略（不影響主流程）

        for r in range(1, total_lines + 1):
            if is_eof:
                break
            line_no = r
            print("=" * 80)
            print(f"處理第 {line_no} 行...")
            row = line_start_row + (r - 1) * rows_per_line + han_ji_row_offset

            new_line = False
            for c in range(start_col, end_col + 1):
                if is_eof:
                    break  # noqa: E701
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
            self.save_all_piau_im_ji_khoo_dicts()


def process(wb, args) -> int:
    """
    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    # ------------------------------------------------------------------------------
    # 作業初始化
    # ------------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        # 初始化 process config
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=True,
            new_piau_im_ji_khoo_sheet=True,
            new_khuat_ji_piau_sheet=True,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    try:
        # ======================================================================
        # 將【漢字注音】工作表的舊資料清除及格式重設。
        # ======================================================================
        # 重置工作表
        print("清除儲存格內容作業...")
        clear_han_ji_kap_piau_im(
            wb,
            sheet_name="漢字注音",
            total_lines=program.TOTAL_LINES,
            rows_per_line=program.ROWS_PER_LINE,
            start_row=program.line_start_row,
            start_col=program.start_col,
            end_col=program.end_col,
            han_ji_orgin_cell=program.han_ji_orgin_cell,
        )
        # logging.info("儲存格內容清除完畢")
    except Exception as e:
        logging_exc_error(msg="清除儲存格內容作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    try:
        if args.reset_cell_format:
            print("重設儲存格格式作業...")
            reset_cells_format_in_sheet(
                wb,
                sheet_name="漢字注音",
                total_lines=program.TOTAL_LINES,
                rows_per_line=program.ROWS_PER_LINE,
                start_row=program.line_start_row,
                start_col=program.start_col,
                end_col=program.end_col,
            )
            # logging.info("儲存格格式重設完畢")
    except Exception as e:
        logging_exc_error(msg="重置儲存格格式作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    try:
        # ======================================================================
        # 填入【漢字】：讀取整篇文章之【漢字】純文字檔案；並填入【漢字注音】工作表。
        # ======================================================================
        text_file_name = args.han_ji_file if args.han_ji_file else "_tmp_p1_han_ji.txt"
        # 顯示正在處理的漢字檔案名稱
        print("=" * 80)
        print(f"正在處理的漢字檔案：{text_file_name}")
        _fill_han_ji_into_sheet(
            wb=wb,
            program=program,
            text_file_name=text_file_name,
            sheet_name="漢字注音",
            target="V3",
        )
    except Exception as e:
        logging_exc_error(msg="將【漢字】填入【漢字注音】工作表異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    try:
        # ==========================================================================
        # 將【漢字注音】工作表的【漢字】欄，逐一處理，查找【台語音標】和【漢字標音】
        # ==========================================================================

        # 處理工作表
        sheet_name = "漢字注音"
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 處理整張工作表的各個儲存格
        xls_cell._process_sheet(sheet)
    except Exception as e:
        logging_exception(
            msg=f"在【{sheet_name}】工作表，自動為【漢字】查找【台語音標】作業，發生例外！",
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
    wb = None
    try:
        # 嘗試從 Excel 呼叫取得（RunPython）
        wb = xw.Book.caller()
    except Exception:
        # 若失敗，則取得作用中的活頁簿
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
            return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        logging.error("無法取得 Excel 活頁簿")
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
        logging.error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 要求畫面回到【漢字注音】工作表
        # wb.sheets['漢字注音'].activate()
        # 儲存檔案
        if not Program.save_workbook_as_new_file(wb=wb):
            return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


def test_01():
    """測試函數能計算總行數"""
    from mod_excel_access import calculate_total_lines

    # 取得 wb 物件以供設定 sheet
    wb = None
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}！")
        return EXIT_CODE_NO_FILE

    sheet = wb.sheets["漢字注音"]
    total_lines = calculate_total_lines(sheet)
    if total_lines is not None:
        # 應回傳 158
        print(f"總漢字注音行數：{total_lines}")
    else:
        print("無法計算總漢字注音行數")
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="將漢字純文字檔中的漢字，填入 Excel 活頁簿中的【漢字注音】工作表，並自動查找台語音標與漢字標音。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  python a100_作業中活頁檔填入漢字.py          # 執行一般模式
  python a100_作業中活頁檔填入漢字.py -new     # 建立新的字庫工作表
  python a100_作業中活頁檔填入漢字.py -test    # 執行測試模式
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
    parser.add_argument(
        "han_ji_file", nargs="?", default="_tmp_p1_han_ji.txt", help="漢字純文字檔路徑"
    )
    parser.add_argument("ping_im_file", nargs="?", default="", help="標音檔（可選）")
    parser.add_argument(
        "--reset_cell_format", action="store_true", help="重置工作表初始狀態"
    )
    parser.add_argument("--peh_ue", action="store_true", help="將語音類型設定為白話音")
    parser.add_argument(
        "--tiau_hu",
        action="store_false",
        dest="tiau_ho",
        help="TLPA音標改【聲調符號】（不帶調號數值）",
    )
    args = parser.parse_args()

    if args.test:
        # 執行測試
        test_01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code != EXIT_CODE_SUCCESS:
            logging.error(f"程式異常終止，返回失敗碼：{exit_code}")
            sys.exit(exit_code)
