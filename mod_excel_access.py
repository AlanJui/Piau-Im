"""
mod_excel_access.py v0.2.2.2
提供 Excel 檔案存取相關的輔助函式
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sys
from pathlib import Path
from typing import Optional

# 載入第三方套件
import win32com.client  # 用於獲取作用中的 Excel 檔案

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

from mod_logging import init_logging, logging_exc_error, logging_process_step

# 載入自訂模組
from mod_piau_im_tng_huan import _has_meaningful_data

# =========================================================================
# 常數定義
# =========================================================================

# --------------------------------------------------------------------------
# 儲存格位置常數
#  - 每 1 【行】，內含 4 row ；第 1 行之 row no 為：3
#  - row 1: 人工標音儲存格 ===> row_no= 3,  7, 11, ...
#  - row 2: 台語音標儲存格 ===> row_no= 4,  8, 12, ...
#  - row 3: 漢字儲存格     ===> row_no= 5,  9, 13, ...
#  - row 4: 漢字標音儲存格 ===> row_no= 6, 10, 14, ...
#
# 依【作用儲存格】的 row no 求得：line_no = ((row_no - start_row_no) // rows_per_line) + 1
#
# 依【line_no】求得【基準列 row no】：base_row_no = start_row_no + ((line_no - 1) * rows_per_line)
# --------------------------------------------------------------------------
ROWS_PER_LINE = 4
START_ROW_NO = 3  # 第 1 行的起始列號
START_COL = 4  # D 欄
END_COL = 18  # R 欄

TAI_GI_IM_PIAU_OFFSET = 1
HAN_JI_OFFSET = 2
HAN_JI_PIAU_IM_OFFSET = 3

# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# 預設應有之工作表
DEFAULT_SHEET_LIST = [
    "漢字注音",
    "缺字表",
    "字庫表",
]

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv("DB_HO_LOK_UE", "Ho_Lok_Ue.db")
DB_KONG_UN = os.getenv("DB_KONG_UN", "Kong_Un.db")

# =========================================================================
# 輔助函式
# =========================================================================


def get_full_path_from_workbook(wb) -> str:
    """
    獲取 Excel 活頁簿的完整路徑及所在目錄
    # wb = xw.Book('您的檔案.xlsx')

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :return: (完整路徑, 所在目錄)
    """

    # 取得完整路徑 (例如: C:\work\Piau-Im\test.xlsx)
    full_path = wb.fullname

    # print(f"Excel 活頁簿檔完整路徑: {full_path}")
    return full_path


def get_current_directory_from_workbook(wb) -> str:
    """
    獲取 Excel 活頁簿的所在目錄
    # wb = xw.Book('您的檔案.xlsx')

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :return: 所在目錄
    """

    # 取得完整路徑 (例如: C:\work\Piau-Im\test.xlsx)
    full_path = wb.fullname

    # 使用 pathlib 取得所在目錄
    current_dir = Path(full_path).parent

    # print(f"Excel 活頁簿檔所在目錄: {current_dir}")
    return str(current_dir)


# 方法 1: 檢查是否為 list 且內容是 tuple
def is_coordinate_list(obj):
    return (
        isinstance(obj, list)
        and len(obj) > 0
        and all(isinstance(item, tuple) and len(item) == 2 for item in obj)
    )


# 方法 2: 更嚴格的檢查（包含型別）
def is_coordinate_list_type(obj):
    return isinstance(obj, list) and all(
        isinstance(item, tuple)
        and len(item) == 2
        and all(isinstance(coord, int) for coord in item)
        for item in obj
    )


# -------------------------------------------------------------------------
# 計算工作表中有效列數
# -------------------------------------------------------------------------
def calculate_total_lines(
    sheet,
    rows_per_line=ROWS_PER_LINE,
    start_row_no=START_ROW_NO,
    han_ji_offset=HAN_JI_OFFSET,
    start_col=START_COL,
    end_col=END_COL,
) -> int:
    """
    計算工作表中【總漢字注音行】數
    說明：
    1. 掃描範圍依據 sheet.used_range 決定。
    2. 採一次性讀取資料至記憶體，提升效能。
    3. 若遇「φ」符號則視為結束，回傳該行；否則回傳最後一個有資料的行號。
    4. 可容許中間有空行，不會因此提早中止。

    Args:
        sheet: 工作表物件
        rows_per_line: 每一行包含的列數 (預設 4)
        start_row_no: 起始行的列號 (預設 3)
        han_ji_offset: 漢字列的偏移量 (預設 2)
        start_col: 起始欄 (預設 4, 即 D欄)
        end_col: 結束欄 (預設 18, 即 R欄)

    Returns:
        int: 最後一個有效行的行號（從 1 開始）
    """
    try:
        # 1. 找出工作表有使用的最後一列
        last_cell = sheet.used_range.last_cell
        max_row = last_cell.row

        if max_row < start_row_no:
            return 0

        # 2. 一次讀取範圍內的資料，提升效能
        #    使用 ndim=2 確保回傳二維列表
        #    讀取範圍： (start_row_no, start_col) 到 (max_row, end_col)
        data = (
            sheet.range((start_row_no, start_col), (max_row, end_col))
            .options(ndim=2)
            .value
        )

        last_valid_line = 0
        total_data_rows = len(data)

        # 3. 遍歷資料檢查
        #    row_idx 為相對於 start_row_no 的位移 (0-based)
        #    以 rows_per_line 為步進值檢查每一「行」
        for row_idx in range(0, total_data_rows, rows_per_line):
            # 計算當前行號 (1-based)
            line_no = (row_idx // rows_per_line) + 1

            # 定位到該行的漢字列索引
            han_ji_idx = row_idx + han_ji_offset

            # 確保索引不超出範圍
            if han_ji_idx >= total_data_rows:
                break

            row_values = data[han_ji_idx]

            # 檢查這一列是否有內容
            has_content = False
            is_end_mark = False

            for val in row_values:
                if val is not None:
                    s_val = str(val).strip()
                    if s_val:
                        has_content = True
                        if s_val == "φ":
                            is_end_mark = True
                        # 只要發現有內容，即可停止檢查此列其餘欄位
                        # 但若要找 φ，需確認是否就是 φ，或是有其他內容
                        # 此處邏輯：找到內容標記為有效；若該內容是 φ 標記為結束
                        break

            if has_content:
                last_valid_line = line_no
                if is_end_mark:
                    # 遇到結束符號，直接回傳當前行號
                    return last_valid_line

        return last_valid_line

    except Exception as e:
        print(f"計算總行數時發生錯誤: {e}")
        return 0


def _col_to_index(col) -> int:
    """將欄位字母（如 'D'）轉成欄位序號（如 4）；已是數字則原樣回傳。"""
    if isinstance(col, int):
        return col
    col_number = 0
    for letter in str(col).strip().upper():
        col_number = col_number * 26 + (ord(letter) - ord("A") + 1)
    return col_number


def calculate_total_rows(
    sheet,
    start_col=START_COL,
    end_col=END_COL,
    base_row=START_ROW_NO,
    rows_per_group=ROWS_PER_LINE,
):
    """Compute how many row groups exist based on the described worksheet layout."""
    total_rows = 0
    current_base = base_row

    # 欄位參數可能為欄位字母（如 "D"）或欄位序號（如 4），一律轉成序號
    start_col_idx = _col_to_index(start_col)
    end_col_idx = _col_to_index(end_col)

    while True:
        han_ji_row = current_base + 2   # 漢字所在 row
        piau_im_row = current_base + 3  # 漢字標音所在 row
        target_range = sheet.range(
            (han_ji_row, start_col_idx), (piau_im_row, end_col_idx)
        )
        values = target_range.value

        if not _has_meaningful_data(values):
            break

        total_rows += 1

        # 漢字列出現【文章終止】符號（φ），即終止統計
        han_ji_values = values[0] if isinstance(values[0], list) else [values[0]]
        if any(
            # val is not None and str(val).strip() == "φ"
            val is not None and (str(val).strip() == "φ" or str(val).strip() == "\n")
            for val in han_ji_values
        ):
            break

        current_base += rows_per_group

    return total_rows


def get_row_col_from_coordinate(coord_str):
    """
    自座標字串 `(row, col)` 取出 row, col 座標數值

    :param coord_str: 例如 "(9, 4)"
    :return: row, col 整數座標： 9, 4
    """
    coord_str = coord_str.strip("()")  # 去除括號
    try:
        row, col = map(int, coord_str.split(", "))
        return int(row), int(col)  # 轉換成整數
    except ValueError:
        return ""  # 避免解析錯誤


# 定義儲存格格式
def set_range_format(range_obj, font_name, font_size, font_color, fill_color=None):
    range_obj.api.Font.Name = font_name
    range_obj.api.Font.Size = font_size
    range_obj.api.Font.Color = font_color
    if fill_color:
        # range_obj.api.Interior.Color = fill_color
        # range_obj.color = (255, 255, 204)  # 淡黃色
        range_obj.color = fill_color
    else:
        # range_obj.api.Interior.Pattern = xw.constants.Pattern.xlPatternNone  # 無填滿
        range_obj.color = None


# --------------------------------------------------------------------------
# 清除儲存格內容
# --------------------------------------------------------------------------
def clear_han_ji_kap_piau_im(
    wb,
    sheet_name: str = "漢字注音",
    total_lines: Optional[int] = 120,
    rows_per_line: Optional[int] = 4,
    start_row: Optional[int] = 3,
    start_col: Optional[int] = 4,
    end_col: Optional[int] = 18,
    han_ji_orgin_cell: Optional[str] = "V3",
):
    """清除【工作表】之儲存格內存信

    Args:
        wb (_type_): _description_
        sheet_name (str, optional): _description_. Defaults to '漢字注音'.
    """
    sheet = wb.sheets[sheet_name]  # 選擇工作表
    # sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
    sheet.select()  # 將「漢字注音」工作表設為作用中工作表

    # 每頁最多處理的列數
    total_lines = int(total_lines)  # 從名稱【每頁總列數】取得值
    rows_per_line = int(rows_per_line)  # 每行佔用的列數

    rows_per_line = 4
    end_of_rows = start_row + (total_lines * rows_per_line) - 1
    start_col_name = xw.utils.col_name(start_col)  # D
    end_col_name = xw.utils.col_name(end_col)  # R
    cells_range = f"{start_col_name}{start_row}:{end_col_name}{end_of_rows}"

    # 顯示目前處理【狀態】
    print(f"清除【{sheet_name}】工作表之儲存格內容，範圍為：{cells_range}。")

    # 清除範圍的內容（xlwings 使用 value = None 或 clear() 方法）
    sheet.range(cells_range).value = None
    # sheet.range(cells_range).clear_formats()  # 清除填滿顏色

    # 清空原始漢字儲存格內容（如果有指定的話）
    if han_ji_orgin_cell:
        try:
            sheet.range(han_ji_orgin_cell).value = ""
        except Exception as ex:
            logging.warning(f"無法清空儲存格 {han_ji_orgin_cell}: {ex}")


# 重置【漢字注音】工作表
def reset_cells_format_in_sheet(
    wb,
    sheet_name: Optional[str] = "漢字注音",
    total_lines: Optional[int] = 120,
    rows_per_line: Optional[int] = 4,
    start_row: Optional[int] = 3,
    start_col: Optional[int] = 4,
    end_col: Optional[int] = 18,
):
    try:
        sheet = wb.sheets[sheet_name]  # 選擇【漢字注音】工作表
        rows_per_line = 4
        end_row = start_row + (total_lines * rows_per_line) - 1

        # 設定起始及結束的【欄】位址
        # start_col = 4  # D 欄
        # end_col = start_col + chars_per_row - 1  # 因為欄位是從 1 開始計數

        # 顯示目前處理【狀態】
        start_col_name = xw.utils.col_name(start_col)  # D
        end_col_name = xw.utils.col_name(end_col)  # R
        print(
            f"重置【{sheet_name}】工作表之儲存格格式，範圍為：{start_col_name}{start_row}:{end_col_name}{end_row}。"
        )

        # 以【區塊】（range）方式設置儲存格格式
        row = start_row
        for line in range(1, total_lines + 1):
            # 判斷是否已經超過結束列位址，若是則跳出迴圈
            if row > end_row:
                break
            # 顯示目前處理【狀態】
            # print(f"重置 {line} 行：【漢字】儲存格位於【 {row} 列 】。")
            print(f"重置【漢字注音】第 {line} 行 】。")

            # 人工標音
            range_人工標音 = sheet.range((row - 2, start_col), (row - 2, end_col))
            range_人工標音.value = None
            set_range_format(
                range_人工標音,
                font_name="Arial",
                font_size=24,
                font_color=0xFF0000,  # 紅色
                fill_color=(255, 255, 204),
            )  # 淡黃色

            # 台語音標
            range_台語音標 = sheet.range((row - 1, start_col), (row - 1, end_col))
            range_台語音標.value = None
            set_range_format(
                range_台語音標,
                font_name="Sitka Text Semibold",
                font_size=24,
                font_color=0xFF9933,
            )  # 橙色

            # 漢字
            range_漢字 = sheet.range((row, start_col), (row, end_col))
            range_漢字.value = None
            set_range_format(
                range_漢字,
                font_name="吳守禮細明台語注音",
                font_size=48,
                font_color=0x000000,
            )  # 黑色

            # 漢字標音
            range_漢字標音 = sheet.range((row + 1, start_col), (row + 1, end_col))
            range_漢字標音.value = None
            set_range_format(
                range_漢字標音, font_name="芫荽 0.94", font_size=26, font_color=0x009900
            )  # 綠色

            # 準備處理下一【行】
            row += rows_per_line
    except Exception as e:
        logging_exc_error("重設【漢字注音】工作表儲存格格式時，發生錯誤：", e)
        return EXIT_CODE_PROCESS_FAILURE

    # 返回【作業正常結束代碼】
    return EXIT_CODE_SUCCESS


# --------------------------------------------------------------------------
# 座標位址轉換函式
# --------------------------------------------------------------------------
def convert_to_excel_address(coord_str: tuple[int, int]) -> str:
    """
    轉換 `(row, col)` 格式為 Excel 座標 (如 `(9, 4)` 轉換為 "D9")

    :param coord_str: 例如 "(9, 4)"
    :return: Excel 座標字串，例如 "D9"
    """
    coord_str = coord_str.strip("()")  # 去除括號
    try:
        row, col = map(int, coord_str.split(", "))
        return f"{chr(64 + col)}{row}"  # 轉換成 Excel 座標
    except ValueError:
        return ""  # 避免解析錯誤


def excel_address_to_row_col(cell_address: str) -> tuple[int, int]:
    """
    將 Excel 儲存格地址 (如 'D9') 轉換為 (row, col) 格式。

    :param cell_address: Excel 儲存格地址 (如 'D9', 'AA15')
    :return: (row, col) 元組，例如 (9, 4)
    """
    match = re.match(
        r"([A-Z]+)(\d+)", cell_address
    )  # 用 regex 拆分字母(列) 和 數字(行)

    if not match:
        raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

    col_letters, row_number = match.groups()

    # 將 Excel 字母列轉換成數字，例如 A -> 1, B -> 2, ..., Z -> 26, AA -> 27
    col_number = 0
    for letter in col_letters:
        col_number = col_number * 26 + (ord(letter) - ord("A") + 1)

    return int(row_number), col_number


def excel_address_to_coordinate(cell_address: str) -> tuple[int, int]:
    """
    將 Excel 儲存格地址 (如 'D9') 轉換為 (row, col) 格式。

    :param cell_address: Excel 儲存格地址 (如 'D9', 'AA15')
    :return: (row, col) 元組，例如 (9, 4)
    """
    match = re.match(
        r"([A-Z]+)(\d+)", cell_address
    )  # 用 regex 拆分字母(列) 和 數字(行)

    if not match:
        raise ValueError(f"無效的 Excel 儲存格地址: {cell_address}")

    col_letters, row_number = match.groups()

    # 將 Excel 字母列轉換成數字，例如 A -> 1, B -> 2, ..., Z -> 26, AA -> 27
    col_number = 0
    for letter in col_letters:
        col_number = col_number * 26 + (ord(letter) - ord("A") + 1)

    return int(row_number), col_number


def convert_coord_str_to_excel_address(coord_str: str) -> str:
    """
    轉換 `(row, col)` 格式為 Excel 座標 (如 `(9, 4)` 轉換為 "D9")

    :param coord_str: 例如 "(9, 4)"
    :return: Excel 座標字串，例如 "D9"
    """
    coord_str = coord_str.strip("()")  # 去除括號
    try:
        row, col = map(int, coord_str.split(", "))
        return f"{chr(64 + col)}{row}"  # 轉換成 Excel 座標
    except ValueError:
        return ""  # 避免解析錯誤


def convert_row_col_to_excel_address(row: int, col: int) -> str:
    """
    將 (row, col) 格式轉換為 Excel 座標 (如 (9, 4) 轉換為 "D9")

    :param row: 行號
    :param col: 列號
    :return: Excel 座標字串，例如 "D9"
    """
    return f"{chr(64 + col)}{row}"  # 轉換成 Excel 座標


def strip_cell(x):
    """轉成字串並去除頭尾空白，若空則回傳 None，但保留換行符 \n"""
    # 可以正確區分空白字符和換行符，從而避免將 \n 誤判為空白
    if x is None:
        return None
    x_str = str(x)
    if x_str.strip() == "" and x_str != "\n":  # 空白但不是換行符
        return None
    return x_str.strip() if x_str != "\n" else "\n"  # 保留換行符


def get_active_excel_file():
    """
    獲取當前作用中的 Excel 檔案路徑。
    如果沒有作用中的 Excel 檔案，返回 None。
    """
    try:
        # 獲取 Excel 應用程式
        excel_app = win32com.client.GetObject(Class="Excel.Application")
        if excel_app is None:
            print("❌ 沒有作用中的 Excel 檔案。")
            return None

        # 獲取作用中的工作簿
        active_workbook = excel_app.ActiveWorkbook
        if active_workbook is None:
            print("❌ 沒有作用中的 Excel 工作簿。")
            return None

        # 獲取檔案路徑
        excel_file = active_workbook.FullName
        print(f"✅ 作用中的 Excel 檔案：{excel_file}")
        return excel_file

    except Exception as e:
        print(f"❌ 獲取作用中的 Excel 檔案失敗: {e}")
        return None


def get_line_no_by_row(
    current_row_no, start_row_no=START_ROW_NO, rows_per_line=ROWS_PER_LINE
):
    """
    根據儲存格的 row 座標，計算其所屬的行號 (line no)。

    :param row: 儲存格的 row 座標 (整數)
    :param base_row: 每頁起始列 (預設為 3)
    :param rows_per_group: 每行佔用的列數 (預設為 4)
    :return: 行號 (line no)，從 1 開始計數
    """
    if current_row_no < start_row_no:
        raise ValueError(
            f"儲存格的 row 列號（{current_row_no}）必須大於等於基準列（{START_ROW_NO}）。"
        )
    line_no = ((current_row_no - start_row_no) // rows_per_line) + 1
    return line_no


def get_row_by_line_no(line_no, start_row_no=START_ROW_NO, rows_per_line=ROWS_PER_LINE):
    """
    根據行號 (line no)，計算其對應的儲存格 row 座標。

    :param line_no: 行號 (從 1 開始計數)
    :param base_row: 每頁起始列 (預設為 3)
    :param rows_per_group: 每行佔用的列數 (預設為 4)
    :return: 對應的儲存格 row 座標 (整數)
    """
    if line_no < 1:
        raise ValueError("行號必須大於等於 1。")
    line_base_row_no = start_row_no + ((line_no - 1) * rows_per_line)
    tai_gi_im_piau_row_no = line_base_row_no + TAI_GI_IM_PIAU_OFFSET
    han_ji_row_no = line_base_row_no + HAN_JI_OFFSET
    han_ji_piau_im_row_no = line_base_row_no + HAN_JI_PIAU_IM_OFFSET
    return line_base_row_no, tai_gi_im_piau_row_no, han_ji_row_no, han_ji_piau_im_row_no


def get_active_cell_address():
    """
    獲取目前作用中的 Excel 儲存格地址 (Active Cell Address)

    :return: 儲存格地址字串，例如 "D9"
    """
    try:
        # 獲取 Excel 應用程式
        excel_app = win32com.client.GetObject(Class="Excel.Application")
        if excel_app is None:
            print("❌ 沒有作用中的 Excel 檔案。")
            return None

        # 獲取作用中的儲存格
        active_cell = excel_app.ActiveCell
        if active_cell is None:
            print("❌ 沒有作用中的 Excel 儲存格。")
            return None

        # 獲取儲存格地址
        cell_address = active_cell.Address.replace("$", "")  # 去掉 "$"
        # print(f"✅ 作用中的儲存格地址：{cell_address}")
        return cell_address

    except Exception as e:
        print(f"❌ 獲取作用中的儲存格地址失敗: {e}")
        return None


def get_active_cell_info(wb):
    """
    取得目前 Excel 作用儲存格的資訊：
    - 作用儲存格的位置 (row, col)
    - 取得【漢字】的值
    - 計算【人工標音】儲存格位置，並取得【人工標音】值

    :param wb: Excel 活頁簿物件
    :return: (sheet_name, han_ji, (row, col), artificial_pronounce, (artificial_row, col))
    """
    active_cell = wb.app.selection  # 取得目前作用中的儲存格
    sheet_name = active_cell.sheet.name  # 取得所在的工作表名稱
    cell_address = active_cell.address.replace(
        "$", ""
    )  # 取得 Excel 格式地址 (去掉 "$")

    row, col = excel_address_to_row_col(cell_address)  # 轉換為 (row, col)

    # 取得【漢字】 (作用儲存格的值)
    han_ji = active_cell.value

    # 計算【人工標音】位置 (row-2, col) 並取得其值
    artificial_row = row - 2
    artificial_cell = wb.sheets[sheet_name].cells(artificial_row, col)
    artificial_pronounce = artificial_cell.value  # 取得人工標音的值

    return sheet_name, han_ji, (row, col), artificial_pronounce, (artificial_row, col)


def get_active_cell(wb):
    """
    獲取目前作用中的 Excel 儲存格 (Active Cell)

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :return: (工作表名稱, 儲存格地址)，如 ("漢字注音", "D9")
    """
    active_cell = wb.app.selection  # 獲取目前作用中的儲存格
    sheet_name = active_cell.sheet.name  # 獲取所在的工作表名稱
    cell_address = active_cell.address.replace(
        "$", ""
    )  # 取得 Excel 格式地址 (去掉 "$")

    return sheet_name, cell_address


def set_active_cell(wb, sheet_name, cell_address):
    """
    設定 Excel 作用儲存格位置。

    :param wb: Excel 活頁簿物件 (xlwings.Book)
    :param sheet_name: 目標工作表名稱 (str)
    :param cell_address: 目標儲存格位址 (如 "F33")
    """
    try:
        sheet = wb.sheets[sheet_name]  # 獲取指定工作表
        sheet.activate()  # 確保工作表為作用中的表單
        sheet.range(cell_address).select()  # 設定作用儲存格
        print(f"✅ 已將作用儲存格設為：{sheet_name} -> {cell_address}")
    except Exception as e:
        print(f"❌ 設定作用儲存格失敗: {e}")


def get_sheet_data(sheet, start_cell):
    """
    從指定工作表讀取資料，並確保返回 2D 列表。
    :param sheet: 工作表物件。
    :param start_cell: 起始儲存格（例如 "A2"）。
    :return: 2D 列表，若無資料則返回空列表。
    """
    data = sheet.range(start_cell).expand("table").value
    if data is None:
        return []
    return data if isinstance(data[0], list) else [data]


# --------------------------------------------------------------------------
# 工作表操作函式
# --------------------------------------------------------------------------
# 依工作表名稱，刪除工作表
def delete_sheet_by_name(wb, sheet_name: str, show_msg: bool = False):
    """
    刪除指定名稱的工作表
    wb: Excel 活頁簿物件
    sheet_name: 要刪除的工作表名稱
    """
    try:
        # 檢查工作表是否存在
        if sheet_name in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[sheet_name]
            sheet.delete()  # 刪除工作表
            if show_msg:
                print(f"已成功刪除工作表：{sheet_name}")
        else:
            if show_msg:
                print(f"無法刪除，工作表 {sheet_name} 不存在")
    except Exception as e:
        if show_msg:
            print(f"刪除工作表時發生錯誤：{e}")


# 使用 List 刪除工作表
def delete_sheets_by_list(wb, sheet_list: list, show_msg: bool = False):
    """
    刪除指定名稱的工作表
    wb: Excel 活頁簿物件
    sheet_list: 要刪除的工作表名稱清單
    """
    for sheet_name in sheet_list:
        delete_sheet_by_name(wb, sheet_name, show_msg)


def ensure_sheet_exists(wb, sheet_name):
    """
    確保指定名稱的工作表存在，如果不存在則新增。

    :param wb: Excel 活頁簿物件。
    :param sheet_name: 工作表名稱。
    :return: 確保存在的工作表物件。
    """
    try:
        # 先確保 `wb` 不是 None，並且 `wb.sheets` 可以被存取
        if not wb or not wb.sheets:
            raise ValueError("Excel 活頁簿 `wb` 無效或未正確開啟！")

        # **使用 `name` 屬性來檢查是否存在該工作表**
        sheet_names = [sheet.name for sheet in wb.sheets]

        if sheet_name in sheet_names:
            sheet = wb.sheets[sheet_name]  # 取得現有工作表
        else:
            sheet = wb.sheets.add(sheet_name)  # 新增工作表

        return sheet

    except Exception as e:
        print(f"⚠️ 無法確保工作表存在: {e}")
        return None  # 若發生錯誤，返回 None


def get_value_by_name(wb, name: str, sheet_name=None):
    """
    從 Excel 活頁簿中取得具名範圍的值
    wb: Excel 活頁簿物件
    name: 具名範圍名稱
    sheet_name: (選用) 指定工作表名稱，若指定則嘗試搜尋該工作表的具名範圍
    """
    try:
        # 1. 嘗試直接從活頁簿全域名稱取值
        if name in wb.names:
            return wb.names[name].refers_to_range.value

        # 2. 若指定工作表，嘗試從工作表取值
        if sheet_name:
            try:
                sheet = wb.sheets[sheet_name]
                if name in sheet.names:
                    return sheet.names[name].refers_to_range.value
            except Exception:
                pass

        # 3. 嘗試遍歷所有工作表 (若名字是區域性的)
        # 注意：這可能會因為不同工作表有同名範圍而取到非預期值，
        # 但通常使用者只會在一個工作表設定這些參數
        for sheet in wb.sheets:
            if name in sheet.names:
                return sheet.names[name].refers_to_range.value

        return None

    except Exception:
        return None


def get_ji_khoo(wb, sheet_name="標音字庫"):
    """
    從 Excel 工作表中取得漢字庫
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    """
    # 取得或新增工作表
    if sheet_name not in [s.name for s in wb.sheets]:
        sheet = wb.sheets.add(sheet_name, after=wb.sheets["漢字注音"])
        print(f"已新增工作表：{sheet_name}")
        # 新增標題列
        sheet.range("A1").value = ["漢字", "台語音標", "總數", "校正音標"]
    else:
        sheet = wb.sheets[sheet_name]

    return sheet


def maintain_ji_khoo(sheet, han_ji, tai_gi, show_msg=False):
    """
    維護【漢字庫】工作表，新增或更新漢字及台語音標
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    han_ji: 要新增的漢字
    tai_gi: 對應的台語音標
    """
    # 如果台語音標為空字串，設置為"NA"（或其他標示值）
    tai_gi = tai_gi if tai_gi.strip() else "NA"

    # 取得 A、B、C 欄的所有值
    data = sheet.range("A2").expand("table").value

    # 如果只有一行資料，將其轉換為 2D 列表
    if data and not isinstance(data[0], list):
        data = [data]

    if data is None:  # 如果工作表中沒有資料
        data = []

    # 將資料轉換為標準的列表格式，並將空白欄位替換為空字串
    records = [[r if r is not None else "" for r in row] for row in data]

    # 檢查是否已存在相同的「漢字」和「台語音標」
    found = False
    for i, row in enumerate(records):
        if row[0] == han_ji and row[1] == tai_gi:
            row[2] = (
                row[2] if isinstance(row[2], (int, float)) else 0
            ) + 1  # 確保存在總數是數字
            found = True
            if show_msg:
                print(f"漢字：【{han_ji}（{tai_gi}）】紀錄己有，總數為： {int(row[2])}")
            break

    # 若未找到則新增一筆資料
    if not found:
        records.append([han_ji, tai_gi, 1])
        if show_msg:
            print(f"新增漢字：【{han_ji}】（{tai_gi}）")

    # 更新工作表的內容
    sheet.range("A2").expand("table").clear_contents()  # 清空舊資料
    sheet.range("A2").value = records  # 寫入更新後的資料

    # if show_msg: print(f"已完成【漢字庫】工作表的更新！")


def get_tai_gi_by_han_ji(sheet, han_ji, show_msg=False):
    """
    根據漢字取得台語音標
    wb: Excel 活頁簿物件
    sheet_name: 工作表名稱
    han_ji: 欲查詢的漢字
    """
    # 取得 A、B 欄的所有值
    data = sheet.range("A2").expand("table").value

    if data is None:  # 如果工作表中沒有資料
        if show_msg:
            print("【漢字庫】工作表中沒有任何資料")
        return None

    # 確保資料為 2D 列表
    if not isinstance(data[0], list):
        data = [data]

    # 將資料轉換為標準格式，並查找對應的台語音標
    for row in data:
        han_ji_cell = row[0] if row[0] is not None else ""
        tai_gi_cell = row[1] if row[1] is not None else ""
        if han_ji_cell == han_ji:
            if show_msg:
                print(f"找到台語音標：【{tai_gi_cell}】")
            return tai_gi_cell

    if show_msg:
        print(f"漢字：【{han_ji}】不存在於【漢字庫】")
    return None


def create_dict_by_piau_im_sheet(
    wb, sheet_name="標音字庫", start_row: int =2, end_col: str = "D"
) -> Optional[dict]:
    """
    以標音用工作表，建置查詢用字典，key: 漢字, value: (台語音標, 校正音標, 次數)
    """
    # 取得工作表
    sheet = wb.sheets[sheet_name]
    sheet.activate()

    try:
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
        if last_row < start_row:
            print("Excel 無資料 (至少需要有一列資料)。")
            return []

        data = sheet.range(f"A{start_row}:{end_col}{last_row}").value
        if not isinstance(data[0], list):
            data = [data]

        dict_list = []
        for row in data:
            dict_list.append({
                "漢字": row[0],
                "台語音標": row[1],
                "校正音標": row[2],
                "座標": row[3]
            })
        return dict_list

    except Exception as e:
        print(f"讀取 Excel 資料失敗: {e}")
        return []


# def create_dict_by_sheet(
#     wb, sheet_name: str, allow_empty_correction: bool = False
# ) -> Optional[dict]:
#     """
#     更新【標音字庫】表中的【台語音標】欄位內容，依據【漢字注音】表中的【人工標音】欄位進行更新，並將【人工標音】覆蓋至原【台語音標】。
#     """
#     # 取得工作表
#     work_sheet = wb.sheets[sheet_name]
#     work_sheet.activate()

#     # 取得【標音字庫】表格範圍的所有資料
#     data = work_sheet.range("A2").expand("table").value

#     if data is None:
#         print(f"【{sheet_name}】工作表無資料")
#         return None

#     # 確保資料為 2D 列表
#     if not isinstance(data[0], list):
#         data = [data]

#     # 將資料轉為字典格式，key: 漢字, value: (台語音標, 校正音標, 次數)
#     han_ji_dict = {}
#     for i, row in enumerate(data, start=2):
#         han_ji = row[0] or ""
#         tai_gi_im_piau = row[1] or ""
#         total_count = (
#             int(row[2]) if len(row) > 2 and isinstance(row[2], (int, float)) else 0
#         )
#         corrected_tai_gi = row[3] if len(row) > 3 else ""  # 若無 D 欄資料則設為空字串

#         # 在 dict 新增一筆紀錄：（1）已填入校正音標，且校正音標不同於現有之台語音標；（2）允許校正音標為空時也加入字典
#         if allow_empty_correction or (
#             corrected_tai_gi and corrected_tai_gi != tai_gi_im_piau
#         ):
#             han_ji_dict[han_ji] = (
#                 tai_gi_im_piau,
#                 corrected_tai_gi,
#                 total_count,
#                 i,
#             )  # i 為資料列索引

#     # 若 han_ji_dict 為空，表查找不到【漢字】對應的【台語音標】
#     if not han_ji_dict:
#         print(f"無法依據【{sheet_name}】工作表，建置【字庫】字典")
#         return None

#     return han_ji_dict


def get_sheet_by_name(wb, sheet_name="工作表1"):
    try:
        # 嘗試取得工作表
        sheet = wb.sheets[sheet_name]
        print(f"取得工作表：{sheet_name}")
    except Exception:
        # 若不存在，則新增工作表
        print(f"無法取得，故新建工作表：{sheet_name}...")
        sheet = wb.sheets.add(sheet_name, after=wb.sheets[-1])
        print(f"新建工作表：{sheet_name}")

    # 傳回 sheet 物件
    return sheet


def prepare_working_sheets(wb, sheet_list=DEFAULT_SHEET_LIST):
    # 確認作業用工作表已存在；若無，則建置
    for sheet_name in sheet_list:
        sheets = [sheet.name for sheet in wb.sheets]  # 獲取所有工作表的名稱
        if sheet_name in sheets:
            sheet = wb.sheets[sheet_name]
            try:
                sheet.select()
                sheet.clear()
                continue
            except Exception as e:
                print(e)
        else:
            # CommandError 的 Exception 發生時，表工作表不存在
            # 新增程式需使用之工作表
            print(f"工作表 {sheet_name} 不存在，正在新增...")
            wb.sheets.add(name=sheet_name)


def get_total_rows_in_sheet(wb, sheet_name):
    try:
        # 取得指定的工作表
        sheet = wb.sheets[sheet_name]

        # 從最後一列往上尋找第一個有資料的儲存格所在的列
        last_row = sheet.range("A1048576").end("up").row

        # 若 A1 也為空，代表整個 A 欄都沒有資料
        if sheet.range(f"A{last_row}").value is None:
            total_rows = 0
        else:
            total_rows = last_row

    except Exception as e:
        print(f"無法取得工作表：{sheet_name} （錯誔訊息：{e}）")
        total_rows = 0

    return total_rows


# =========================================================================
# 單元測試
# =========================================================================
def ut_get_sheet_data(wb=None):
    if not wb:
        wb = xw.Book("Test_Case_Sample.xlsx")
    sheet = wb.sheets["漢字注音"]
    data = get_sheet_data(sheet, "D5")
    for row in data:
        print(row)
    return EXIT_CODE_SUCCESS


def ut_khuat_ji_piau(wb=None):
    """缺字表登錄單元測試"""
    wb = xw.Book("Test_Case_Sample.xlsx")
    wb.activate()
    delete_sheet_by_name(wb, "缺字表", show_msg=True)
    sheet = get_ji_khoo(wb, "缺字表")
    sheet.activate()

    try:
        # 當【缺字表】工作表，尚不存在任何查找不到【標音】的【漢字】，新增一筆紀錄
        maintain_ji_khoo(sheet, "銜", "", show_msg=True)
        # 當【缺字表】已有一筆紀錄，新增第二筆紀錄
        maintain_ji_khoo(sheet, "暉", "", show_msg=True)
        # 在【缺字表】新增第三紀錄
        maintain_ji_khoo(sheet, "霪", "", show_msg=True)
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    # 檢查【缺字表】工作表的內容
    for row in sheet.range("A2").expand("table").value:
        print(row)
    return EXIT_CODE_SUCCESS


def ut_maintain_han_ji_koo(wb=None):
    wb = xw.Book("Test_Case_Sample.xlsx")
    sheet = get_ji_khoo(wb, "漢字庫")

    # 漢字庫工作表不存在：工作表將新增，且新增一筆紀錄，加入【說】字，【總數】為 1
    maintain_ji_khoo(sheet, "說", "sue3", show_msg=True)
    # 再次要求在漢字庫加入【說】：工作表會被選取，不會為【說】添增新紀錄，但【總數】更新為 2
    maintain_ji_khoo(sheet, "說", "sue3", show_msg=True)
    maintain_ji_khoo(sheet, "說", "sue3", show_msg=True)
    maintain_ji_khoo(sheet, "說", "uat4", show_msg=True)
    maintain_ji_khoo(sheet, "花", "hua1", show_msg=True)
    maintain_ji_khoo(sheet, "說", "uat4", show_msg=True)

    # 查詢【漢字】的台語音標
    print("\n===================================================")
    han_ji = "說"
    tai_gi = get_tai_gi_by_han_ji(sheet, han_ji)
    if tai_gi:
        print(f"查到【{han_ji}】的台語音標為：{tai_gi}")
    else:
        print(f"查不到【{han_ji}】的台語音標！")

    print("\n===================================================")
    han_ji = "龓"
    tai_gi = get_tai_gi_by_han_ji(sheet, han_ji)
    if tai_gi:
        print(f"查到【{han_ji}】的台語音標為：{tai_gi}")
    else:
        print(f"查不到【{han_ji}】的台語音標！")

    return EXIT_CODE_SUCCESS


def ut_prepare_working_sheets(wb=None):
    if not wb:
        wb = xw.Book()

    #  工作表已存在
    try:
        prepare_working_sheets(wb)
        print("工作表已存在")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表不存在
    try:
        prepare_working_sheets(wb, sheet_list=["工作表1", "工作表2"])
        print("工作表不存在")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS


def ut_get_sheet_by_name(wb=None):
    if not wb:
        wb = xw.Book()

    #  工作表已存在
    try:
        sheet = get_sheet_by_name(wb, "漢字注音")
        print(sheet.name)
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表不存在
    try:
        sheet = get_sheet_by_name(wb, "字庫表")
        print(sheet.name)
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS


def ut10_get_total_rows_in_sheet(wb=None, sheet_name="字庫表") -> int:
    #  工作表已存在
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表無資料
    sheet_name = "工作表1"
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    #  工作表不存在
    sheet_name = "X"
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS


def ut20_get_total_rows_in_sheet(wb=None, sheet_name="漢字注音") -> int:
    #  工作表已存在
    try:
        total_rows = get_total_rows_in_sheet(wb, sheet_name)
        print(f"工作表 {sheet_name} 共有 {total_rows} 列")
    except Exception as e:
        print(e)
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS


def ut01_取得當前作用儲存格(wb):
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    # 取得作用中儲存格的值
    active_cell = wb.sheets[sheet_name].range(cell_address)
    cell_value = active_cell.value
    print(f"📌 作用儲存格{cell_address}的值為：{cell_value}")

    # 將 (row, col) 格式轉換為 Excel 儲存格地址
    # new_cell_address = convert_to_excel_address(cell_value)
    new_cell_address = convert_to_excel_address(f"({row}, {col})")
    print(f"📌 {cell_value} 座標，其 Excel 位址為：{new_cell_address}")

    # 利用 Excel 儲存格地址，將【標音字庫】工作表的 Excel 儲存格位置設為作用儲存格
    target_sheet = "漢字注音"
    target_cell_address = new_cell_address
    set_active_cell(wb, target_sheet, target_cell_address)

    return EXIT_CODE_SUCCESS


def ut02_利用列欄座標值定位漢字注音儲存格(wb):
    sheet_name = "人工標音字庫"
    cell_address = "E2"
    set_active_cell(wb, sheet_name, cell_address)

    # 取得作用中儲存格的值
    active_cell = wb.sheets[sheet_name].range(cell_address)
    cell_value = active_cell.value
    print(f"📌 作用儲存格{cell_address}的值為：{cell_value}")

    # 將 Excel 儲存格地址 "C2" (2, 3) ，轉換成 (row, col) 座標格式
    new_cell_address = convert_to_excel_address(cell_address=cell_value)
    print(f"📌 {cell_value} 座標，其 Excel 位址為：{new_cell_address}")

    # 利用 Excel 儲存格地址，將【標音字庫】工作表的 Excel 儲存格位置設為作用儲存格
    target_sheet = "漢字注音"
    target_cell_address = new_cell_address
    set_active_cell(wb, target_sheet, target_cell_address)

    return EXIT_CODE_SUCCESS


def ut99_calculate_total_lines(wb) -> int:
    """計算【漢字注音】工作表的【漢字注音行】總行數單元測試"""
    # from mod_excel_access import calculate_total_lines
    # 假設 wb 是您的活頁簿物件
    sheet = wb.sheets["漢字注音"]
    total_lines = calculate_total_lines(sheet)
    if total_lines is not None:
        # 應回傳 158
        print(f"總漢字注音行數：{total_lines}")
    else:
        print("無法計算總漢字注音行數")
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS


# =========================================================================
# 作業程序
# =========================================================================
def process(wb) -> int:
    # ---------------------------------------------------------------------
    total_lines = ut99_calculate_total_lines(wb)
    if total_lines is not None:
        # 應回傳 158
        print(f"總漢字注音行數：{total_lines}")
    else:
        print("無法計算總漢字注音行數")
        return EXIT_CODE_UNKNOWN_ERROR

    return EXIT_CODE_SUCCESS
    # ---------------------------------------------------------------------
    # return_code = ut02_利用列欄座標值定位漢字注音儲存格(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut01_取得當前作用儲存格(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_get_sheet_data(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_khuat_ji_piau(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # han_ji_dict = create_dict_by_sheet(wb=wb, sheet_name='缺字表', allow_empty_correction=True)
    # han_ji = '霪'
    # if han_ji_dict and han_ji in han_ji_dict:
    #     original_tai_gi, corrected_tai_gi, total_count, row_index_in_ji_khoo = han_ji_dict[han_ji]
    #     if not corrected_tai_gi:
    #         corrected_tai_gi = "NA"
    #     print(f"【{han_ji}】的台語音標為：{original_tai_gi}，校正音標為：{corrected_tai_gi}，總數：{total_count}，列索引：{row_index_in_ji_khoo}")
    # else:
    #     return EXIT_CODE_PROCESS_FAILURE
    # ---------------------------------------------------------------------
    # return_code = ut_maintain_han_ji_koo(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_prepare_working_sheets(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_get_sheet_by_name(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------
    # return_code = ut_get_total_rows_in_sheet(wb=wb)
    # if return_code != EXIT_CODE_SUCCESS:
    #     return return_code
    # ---------------------------------------------------------------------

    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    print(f"專案根目錄為: {project_root}")
    logging.info(f"專案根目錄為: {project_root}")

    # =========================================================================
    # (2) 設定【作用中活頁簿】：偵測及獲取 Excel 已開啟之活頁簿檔案。
    # =========================================================================
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        print(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("作業異常終止！")
            return result_code

    except Exception as e:
        print(f"作業過程發生未知的異常錯誤: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            print("程式已執行完畢！")

    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
        sys.exit(exit_code)
