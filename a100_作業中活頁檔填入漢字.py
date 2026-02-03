"""
a100_作業中活頁檔填入漢字.py V0.2.6
功能：將漢字純文字檔中的漢字，填入 Excel 活頁簿中的【漢字注音】工作表，並自動查找台語音標與漢字標音。
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


def _process_sheet(sheet, program: Program, xls_cell: ExcelCell) -> None:
    """處理整個工作表"""

    # 處理所有的儲存格
    active_cell = sheet.range(
        f"{xw.utils.col_name(program.start_col)}{program.line_start_row}"
    )
    active_cell.select()

    # 調整 row 值至【漢字】列（每 4 列為一組【列群】，漢字在第 3 列：5, 9, 13, ... ）
    is_eof = False
    for r in range(1, program.TOTAL_LINES + 1):
        if is_eof:
            break
        line_no = r
        print("=" * 80)
        print(f"處理第 {line_no} 行...")
        row = (
            program.line_start_row
            + (r - 1) * program.ROWS_PER_LINE
            + program.han_ji_row_offset
        )
        for c in range(program.start_col, program.end_col + 1):
            if is_eof:
                break  # noqa: E701
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
            status_code = xls_cell._process_cell(active_cell, row, col)

            # 檢查是否需因：換行、文章終結，而跳出內層迴圈
            if status_code == 1:
                is_eof = True
                break
            elif status_code == 2:
                break


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

    text = ""
    for han_ji_ku in text_with_han_ji:
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


def process(wb, args) -> int:
    """
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
        # xls_cell = ExcelCell(program=program)
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=True if args.new else False,
            new_piau_im_ji_khoo_sheet=True if args.new else False,
            new_khuat_ji_piau_sheet=True if args.new else False,
        )

        # ======================================================================
        # 將【漢字注音】工作表的舊資料清除及格式重設。
        # ======================================================================
        # 重置工作表
        print("清除儲存格內容...")
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
        logging.info("儲存格內容清除完畢")

        if args.reset_cell_format:
            print("重設儲存格之格式...")
            reset_cells_format_in_sheet(
                wb,
                sheet_name="漢字注音",
                total_lines=program.TOTAL_LINES,
                rows_per_line=program.ROWS_PER_LINE,
                start_row=program.line_start_row,
                start_col=program.start_col,
                end_col=program.end_col,
            )
            logging.info("儲存格格式重設完畢")

        # ======================================================================
        # 填入【漢字】：讀取整篇文章之【漢字】純文字檔案；並填入【漢字注音】工作表。
        # ======================================================================
        text_file_name = args.han_ji_file if args.han_ji_file else "_tmp_p1_han_ji.txt"
        _fill_han_ji_into_sheet(
            wb=wb,
            program=program,
            text_file_name=text_file_name,
            sheet_name="漢字注音",
            target="V3",
        )

        # ======================================================================
        # 將【漢字注音】工作表的【漢字】欄，逐一處理，查找【台語音標】和【漢字標音】
        # ======================================================================

        # 處理工作表
        sheet_name = "漢字注音"
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 逐列處理
        _process_sheet(
            sheet=sheet,
            program=program,
            xls_cell=xls_cell,
        )

        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dicts()

        print("=" * 80)
        logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS

    except Exception:
        logging.exception("自動為【漢字】查找【台語音標】作業，發生例外！")
        raise


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
    try:
        # 取得 Excel 活頁簿
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

    except Exception as e:
        logging.exception(f"取得作用中活頁簿作業失敗： {e} ！")
        return EXIT_CODE_UNKNOWN_ERROR

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
    pass


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
