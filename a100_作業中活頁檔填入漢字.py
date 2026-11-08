"""
a100_作業中活頁檔填入漢字.py V0.2.14
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
 v0.2.10 2026-10-19:
  - 預設 --new：先移除作用活頁簿的【缺字表】、【標音字庫】、【人工標音字庫】，再建立新表。
  - 加入 --update：略過上述三張工作表的移除，沿用既有工作表。
 v0.2.11 2026-10-27:
  - 抄入漢字之前，先將【漢字】儲存格（列 5、9、13、17…，欄 D:R）底色改回無填滿。
 v0.2.12 2026-11-08:
  - 找不到作用中活頁簿時，改自專案根目錄開啟 Template.xlsx，不再直接終止程式。
  - 清除舊殘留時，一併移除【漢字注音】標音區塊內每個儲存格的附註。
 v0.2.13 2026-11-08:
  - 找不到【漢字注音】工作表時，自【漢字注音【模版】】複製一張再繼續作業。
 v0.2.14 2026-11-08:
  - 清除舊內容、附註與漢字底色時，範圍改涵蓋工作表上已有資料的列，
    不再只清到名稱【每頁總列數】為止。
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
    logging_warning,
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

HAN_JI_SHEET_NAME = "漢字注音"
HAN_JI_TEMPLATE_SHEET_NAME = "漢字注音【模版】"

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()


# =========================================================================
# 主要處理函數
# =========================================================================
def annotation_clear_line_count(sheet, program: Program) -> int:
    """計算要清除的行數。

    名稱【每頁總列數】只代表一頁的編排列數。模版或上一篇若比它更長，
    那些列上的舊漢字、標音與附註也要一併清掉。
    """
    configured = int(program.TOTAL_LINES)
    start_row = program.line_start_row
    rows_per_line = program.ROWS_PER_LINE
    if configured < 1:
        configured = 1

    last_row = start_row
    try:
        used_last = int(sheet.used_range.last_cell.row)
    except Exception as e:
        logging_warning(f"無法讀取工作表已使用範圍，改以【每頁總列數】清除：{e}")
        return configured

    if used_last > start_row:
        block = sheet.range((start_row, program.start_col), (used_last, program.end_col)).value
        if block is not None:
            if not isinstance(block, list):
                block = [block]
            if block and not isinstance(block[0], list):
                block = [block]
            for offset, row_vals in enumerate(block):
                if row_vals is None:
                    continue
                if not isinstance(row_vals, list):
                    row_vals = [row_vals]
                if any(value not in (None, "") for value in row_vals):
                    last_row = start_row + offset

    span = max(last_row - start_row + 1, rows_per_line)
    used_lines = (span + rows_per_line - 1) // rows_per_line
    return max(configured, used_lines)


def clear_han_ji_cell_fill(
    wb,
    program: Program,
    sheet_name: str = "漢字注音",
    total_lines: int | None = None,
) -> None:
    """抄入漢字之前，將【漢字】儲存格底色改回無填滿。

    漢字列為 5、9、13、17…（自起始列起每隔 4 列），欄位為 D:R。
    只清底色，不改儲存格內容與字型。
    """
    sheet = wb.sheets[sheet_name]
    start_row = program.han_ji_start_row
    rows_per_line = program.ROWS_PER_LINE
    start_col = program.start_col
    end_col = program.end_col
    total_lines = int(program.TOTAL_LINES if total_lines is None else total_lines)

    if total_lines < 1:
        return

    start_col_name = xw.utils.col_name(start_col)
    end_col_name = xw.utils.col_name(end_col)
    last_han_ji_row = start_row + (total_lines - 1) * rows_per_line
    logging_process_step(
        "將【漢字】儲存格底色改回無填滿："
        f"{start_col_name}{start_row}:{end_col_name}{last_han_ji_row}"
        f"（列 {start_row} 起，每隔 {rows_per_line} 列）。"
    )

    for line in range(total_lines):
        row = start_row + line * rows_per_line
        sheet.range((row, start_col), (row, end_col)).color = None


def clear_cell_notes(
    wb,
    program: Program,
    sheet_name: str = "漢字注音",
    total_lines: int | None = None,
) -> None:
    """移除【漢字注音】標音區塊內，每個儲存格的【附註】。

    範圍與清除舊內容相同：自起始列起，涵蓋人工標音、台語音標、漢字、漢字標音。
    Excel 的附註即舊版 Comment（Note）。先刪工作表附註集合中落在範圍內的項目，
    再對整段範圍呼叫 ClearComments，避免只清到其中一種。
    """
    sheet = wb.sheets[sheet_name]
    start_row = program.line_start_row
    rows_per_line = program.ROWS_PER_LINE
    start_col = program.start_col
    end_col = program.end_col
    total_lines = int(program.TOTAL_LINES if total_lines is None else total_lines)
    if total_lines < 1:
        return

    end_row = start_row + (total_lines * rows_per_line) - 1
    start_col_name = xw.utils.col_name(start_col)
    end_col_name = xw.utils.col_name(end_col)
    logging_process_step(
        "移除【漢字注音】儲存格附註："
        f"{start_col_name}{start_row}:{end_col_name}{end_row}。"
    )

    removed = _delete_notes_in_range(
        sheet,
        start_row=start_row,
        end_row=end_row,
        start_col=start_col,
        end_col=end_col,
    )
    sheet.range((start_row, start_col), (end_row, end_col)).api.ClearComments()
    logging_process_step(
        f"已清除【{sheet_name}】工作表 {start_col_name}{start_row}:{end_col_name}{end_row} 的儲存格附註"
        f"（附註集合刪除 {removed} 則）。"
    )


def _delete_notes_in_range(
    sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> int:
    """刪除範圍內既有附註。讀不到附註集合時，交由呼叫端以範圍一次清除。"""
    try:
        comments = sheet.api.Comments
        count = int(comments.Count)
    except Exception as e:
        logging_warning(f"無法讀取工作表附註集合，改以範圍清除：{e}")
        return 0

    removed = 0
    for index in range(count, 0, -1):
        try:
            comment = comments.Item(index)
            anchor = comment.Parent
            row = int(anchor.Row)
            col = int(anchor.Column)
            if not (start_row <= row <= end_row and start_col <= col <= end_col):
                continue
            comment.Delete()
            removed += 1
        except Exception as e:
            logging_warning(f"無法刪除第 {index} 則儲存格附註：{e}")
    return removed


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
        # for han_ji in han_ji_ku:
        #     # han_ji_list.append(han_ji)
        han_ji_list.extend(han_ji_ku)
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
        except Exception:  # noqa: BLE001, S110
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
                    break
                if new_line:
                    break  # 跳出內層迴圈，進入下一行處理
                # row = row
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


def ensure_han_ji_sheet(
    wb,
    sheet_name: str = HAN_JI_SHEET_NAME,
    template_name: str = HAN_JI_TEMPLATE_SHEET_NAME,
):
    """取得【漢字注音】。活頁簿沒有這張表時，自【漢字注音【模版】】複製。"""
    sheet_names = [sheet.name for sheet in wb.sheets]
    if sheet_name in sheet_names:
        return wb.sheets[sheet_name]
    if template_name not in sheet_names:
        logging.error(
            f"找不到工作表【{sheet_name}】，也沒有可複製的【{template_name}】。"
            f"現有工作表：{sheet_names}"
        )
        return None

    logging_process_step(f"找不到工作表【{sheet_name}】，自【{template_name}】複製。")
    app = wb.app
    previous_alerts = app.display_alerts
    app.display_alerts = False
    try:
        template = wb.sheets[template_name]
        target = template.copy(after=template, name=sheet_name)
    finally:
        app.display_alerts = previous_alerts

    logging_process_step(f"已自【{template_name}】複製製作【{target.name}】。")
    return target


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
        if ensure_han_ji_sheet(wb) is None:
            return EXIT_CODE_PROCESS_FAILURE
    except Exception as e:
        logging_exc_error(msg="無法準備【漢字注音】工作表！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    try:
        # 初始化 process config
        program = Program(wb, args, hanji_piau_im_sheet_name="漢字注音")

        # 預設 --new：先移除三張字庫工作表再建立。
        # --update：略過移除，沿用作用活頁簿中既有的工作表。
        rebuild_ji_khoo_sheets = getattr(args, "ji_khoo_mode", "new") != "update"
        if rebuild_ji_khoo_sheets:
            logging_process_step("【--new】先移除【缺字表】、【標音字庫】、【人工標音字庫】，再建立新表。")
        else:
            logging_process_step("【--update】略過移除【缺字表】、【標音字庫】、【人工標音字庫】。")

        # 建立儲存格處理器
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=rebuild_ji_khoo_sheets,
            new_piau_im_ji_khoo_sheet=rebuild_ji_khoo_sheets,
            new_khuat_ji_piau_sheet=rebuild_ji_khoo_sheets,
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
        han_ji_sheet = wb.sheets["漢字注音"]
        clear_lines = annotation_clear_line_count(han_ji_sheet, program)
        logging_process_step(
            f"清除舊資料共 {clear_lines} 行"
            f"（名稱【每頁總列數】為 {int(program.TOTAL_LINES)}）。"
        )
        clear_han_ji_kap_piau_im(
            wb,
            sheet_name="漢字注音",
            total_lines=clear_lines,
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
        print("移除儲存格附註作業...")
        clear_cell_notes(wb, program, sheet_name="漢字注音", total_lines=clear_lines)
    except Exception as e:
        logging_exc_error(msg="移除儲存格附註作業異常！", error=e)
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
        # 漢字尚未抄入前，先把漢字列（5、9、13、17…，欄 D:R）底色改回無填滿
        clear_han_ji_cell_fill(wb, program, sheet_name="漢字注音", total_lines=clear_lines)
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
def _open_template_workbook(project_root: Path):
    """找不到作用中活頁簿時，自專案根目錄開啟 Template.xlsx。"""
    file_path = project_root / "Template.xlsx"
    if not file_path.is_file():
        logging.error(f"找不到範本活頁簿：{file_path}")
        return None
    logging_process_step(f"找不到作用中活頁簿，改自專案根目錄開啟範本：{file_path}")
    try:
        return xw.Book(str(file_path))
    except Exception as e:
        logging_exc_error(msg="無法開啟範本活頁簿 Template.xlsx", error=e)
        return None


def _acquire_workbook(project_root: Path):
    """取得作用中活頁簿。沒有已開啟的活頁簿時，改開專案根目錄的 Template.xlsx。"""
    try:
        wb = xw.Book.caller()
        if wb:
            return wb
    except Exception:
        pass

    try:
        app = xw.apps.active
        if app is not None and app.books.count > 0:
            wb = app.books.active
            if wb:
                logging_process_step(f"作用中活頁簿：{wb.name}")
                return wb
    except Exception as e:
        logging_process_step(f"找不到作用中的 Excel 活頁簿：{e}")

    return _open_template_workbook(project_root)


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
    #     找不到時，改自專案根目錄開啟 Template.xlsx。
    # =========================================================================
    wb = _acquire_workbook(project_root)
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
  python a100_作業中活頁檔填入漢字.py              # 預設 --new：先移除三張字庫工作表再建立
  python a100_作業中活頁檔填入漢字.py --new        # 同上
  python a100_作業中活頁檔填入漢字.py --update     # 略過移除，沿用既有字庫工作表
  python a100_作業中活頁檔填入漢字.py --test       # 執行測試模式

找不到作用中活頁簿時，改開專案根目錄的 Template.xlsx。
找不到【漢字注音】工作表時，自【漢字注音【模版】】複製一張。
抄入漢字前，會清除【漢字注音】標音區塊內每個儲存格的附註。
""",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )
    ji_khoo_mode = parser.add_mutually_exclusive_group()
    ji_khoo_mode.add_argument(
        "--new",
        dest="ji_khoo_mode",
        action="store_const",
        const="new",
        help="先移除【缺字表】、【標音字庫】、【人工標音字庫】，再建立新表（預設）",
    )
    ji_khoo_mode.add_argument(
        "--update",
        dest="ji_khoo_mode",
        action="store_const",
        const="update",
        help="略過移除【缺字表】、【標音字庫】、【人工標音字庫】，沿用既有工作表",
    )
    parser.set_defaults(ji_khoo_mode="new")
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
