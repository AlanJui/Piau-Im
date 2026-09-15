"""
a109_漢字注音工作表校對.py V0.3
# =========================================================================
# 程式功能摘要
# =========================================================================
# 用途：提供 <-- 及 --> （向前/向後）按鍵，以利操作者在誦讀【漢字注音】工作表時，
# 可利用【作用儲存格】十字交叉的聚焦游標，導引觀眾目光的移動，使逐字的漢字發音誦讀，
# 更顯有趣。另外，操作者無需借助滑鼠指標，僅需使用【←】或【→】按鍵，便能在上/下行
# 移動。譬如：在【第2行】的行尾（即儲存格：R9）時，按【→】鍵，游標會跳到【第3行】
# 的行首（即儲存格：D13）。
======================================================================
漢字注音工作表導讀（鍵盤監聽模式）
======================================================================
操作說明：
  ← (Left Arrow)  : 向左移動
  → (Right Arrow) : 向右移動
  ↑ (Up Arrow)    : 向上移動到上一行
  ↓ (Down Arrow)  : 向下移動到下一行
  PgUp (Page Up)  : 翻到上一頁
  PgDn (Page Down): 翻到下一頁
  空白            : 查字典更換漢字讀音
  J 鍵            : 查字典指定漢字讀音
  E 鍵            : 手動輸入人工標音
  = 鍵            : 填入人工標音標記
  ESC             : 結束程式
======================================================================

變更紀錄：
V0.5 (2026-07-16): 新增 PgUp/PgDn 按鍵支援：以 Excel 視窗目前可視列數換算
每頁行數，令游標一次向上／向下翻一頁（頁首／頁尾自動截止於第一行／最後一行）。

V0.3 (2026-02-28): 變更【人工標音作業】功能（按鍵：【E】），原先以 a224 程式
要求使用者手動輸入漢字讀音；現在改用 a260 程式，先在《個人字典》查找漢字讀音；
故使用者除了自行手動輸入外；亦可直接套用《個人字典》查得的漢字讀音。

V0.4 (2026-07-14): 重構功能，不可支援外部字典查詢。查字典時，亦可人工輸入漢字
讀音。a250 將字典查得讀音，替換【漢字標音】工作表登錄之資料紀錄；a260 則用於
為某一漢字指定漢字讀音（標注於【人工標音】儲存格）

V0.6 (2026-09-13): 修正空白鍵／J 鍵查字失敗。a250／a260 初始化時讀取命名範圍
會把 Excel 作用儲存格帶到第 1 列，導致「列號必須大於等於基準列（3）」；現改為
由 a300 明確傳入目前漢字儲存格位址，並避免查字失敗後整支導航程式被 COM 錯誤中止。

V0.7 (2026-09-13): 查字進入 input() 前，把 Windows 焦點從 Excel 搶回啟動本程式
的終端機（含 WezTerm）。先前只在呼叫 a250／a260 前切換一次，初始化讀取 Excel
時焦點又被帶走，使用者必須改用滑鼠點 Terminal。

V0.8 (2026-09-15): 新增 --start 參數，Esc 中斷後可自指定儲存格續校。
例如：python a300_漢字注音工作表校對.py --start d133
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import argparse
import logging
import os
import re
import subprocess
import sys
import time

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

try:
    from pynput import keyboard

    HAS_PYNPUT = True
except ImportError:
    HAS_PYNPUT = False
    print("警告：未安裝 pynput 套件，將使用輸入模式")
    print("可執行：pip install pynput")

# Windows API（用於視窗切換）
try:
    import win32con
    import win32gui

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("警告：未安裝 pywin32 套件，可能無法自動切換視窗")
    print("可執行：pip install pywin32")

# COM 執行緒初始化（用於多執行緒環境）
try:
    # import pythoncom

    HAS_PYTHONCOM = True
except ImportError:
    HAS_PYTHONCOM = False

# # 載入 a222 的核心查詢功能（個人字典）
# try:
#     from a222_依作用儲存格在個人字典查找漢字讀音 import (
#         process as ca_han_ji_thak_im_a222,
#     )
#
#     HAS_A222 = True
# except ImportError as e:
#     HAS_A222 = False
#     print(f"警告：無法載入 a222 模組：{e}")

# # 載入 a220 的核心查詢功能（萌典）
# try:
#     from a220_作用儲存格查找萌典漢字讀音 import process as ca_han_ji_thak_im_a220
#
#     HAS_A220 = True
# except ImportError as e:
#     HAS_A220 = False
#     print(f"警告：無法載入 a220 模組：{e}")

# 載入 a224 的核心查詢功能（引用既有標音）
try:
    from a224_引用既有的漢字標音 import (
        process as jin_kang_piau_im_ca_taigi_im_piau,
    )

    HAS_A224 = True
except ImportError as e:
    HAS_A224 = False
    print(f"警告：無法載入 a224 模組：{e}")

# 載入 a250 的核心查詢功能（變更【漢字標音】）
try:
    from a250_更換漢字讀音 import (
        process as ca_han_ji_thok_im_a250,
    )

    HAS_A250 = True
except ImportError as e:
    HAS_A250 = False
    print(f"警告：無法載入 a250 模組：{e}")

# 載入 a260 的核心查詢功能：可先在《個人字典》查找漢字讀音；或手動輸入漢字讀音
try:
    from a260_為單一漢字指定讀音 import (
        process as ca_ji_tian_au_thiam_jin_kang_piau_im,
    )

    HAS_A260 = True
except ImportError as e:
    HAS_A260 = False
    print(f"警告：無法載入 a260 模組：{e}")

from mod_window_focus import activate_console_window, capture_console_hwnd

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_ERROR = 10
EXIT_CODE_UNKNOWN_ERROR = 99

# 工作表設定
SHEET_NAME = "漢字注音"
START_ROW = 5  # 第一行的起始列號
START_COL = 4  # D 欄（第 4 欄）
END_COL = 18  # R 欄（第 18 欄）
ROWS_PER_LINE = 4  # 每行佔用 4 列

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


# =========================================================================
# 核心功能函數
# =========================================================================
def get_line_number(row: int) -> int:
    """
    根據列號計算行號

    Args:
        row: Excel 列號

    Returns:
        行號（1-based）
    """
    # 計算從起始列開始的偏移
    offset = row - START_ROW
    # 每 4 列為一行
    line_no = (offset // ROWS_PER_LINE) + 1
    return line_no


def get_row_from_line(line_no: int) -> int:
    """
    根據行號計算該行的漢字儲存格列號

    Args:
        line_no: 行號（1-based）

    Returns:
        該行漢字儲存格的列號
    """
    return START_ROW + (line_no - 1) * ROWS_PER_LINE


def snap_to_han_ji_row(row: int) -> int:
    """
    將列號對齊到所屬【行】的漢字列。

    每一行佔 4 列：人工標音、台語音標、漢字、漢字標音。
    第一行漢字列為 START_ROW（D5 所在列）。
    """
    block_start_row = START_ROW - 2  # 第一行【人工標音】列
    if row < block_start_row:
        return START_ROW
    block_index = (row - block_start_row) // ROWS_PER_LINE
    return START_ROW + block_index * ROWS_PER_LINE


def parse_start_cell_address(cell_address: str) -> tuple[int, int]:
    """
    將起始儲存格位址（如 d133、D133、$D$133）解析為 (row, col)。

    若列號落在同一【行】的四列區塊內，會對齊到該行的漢字列。
    欄位必須在 D 至 R 之間。
    """
    if not cell_address or not str(cell_address).strip():
        raise ValueError("起始儲存格不可為空白")

    normalized = str(cell_address).strip().upper().replace("$", "")
    match = re.match(r"^([A-Z]+)(\d+)$", normalized)
    if not match:
        raise ValueError(f"無效的儲存格位址：{cell_address}（請使用如 D133 的格式）")

    col_letters, row_text = match.groups()
    col_number = 0
    for letter in col_letters:
        col_number = col_number * 26 + (ord(letter) - ord("A") + 1)
    row_number = int(row_text)

    if col_number < START_COL or col_number > END_COL:
        start_letter = xw.utils.col_name(START_COL)
        end_letter = xw.utils.col_name(END_COL)
        raise ValueError(f"起始儲存格欄位必須在 {start_letter} 至 {end_letter} 之間：{normalized}")

    return snap_to_han_ji_row(row_number), col_number


def argparse_start_cell(value: str) -> str:
    """argparse 型別：先驗證儲存格位址，再保留原始字串。"""
    try:
        parse_start_cell_address(value)
    except ValueError as e:
        raise argparse.ArgumentTypeError(str(e)) from e
    return str(value).strip()


def resolve_start_cell(start_address: str | None, total_lines: int) -> tuple[int, int]:
    """
    決定校對起始儲存格。未指定時自 D5 開始；超出最後一行時改從最後一行開始。
    """
    if not start_address:
        return START_ROW, START_COL

    row, col = parse_start_cell_address(start_address)
    last_han_ji_row = get_row_from_line(max(1, total_lines))
    if row > last_han_ji_row:
        print(f"⚠️  起始列超出最後一行（漢字列 {last_han_ji_row}），改從最後一行開始")
        row = last_han_ji_row
    return row, col


def move_up(sheet, current_row: int, current_col: int) -> tuple:
    """
    向上移動游標到上一行的相同欄位（或行首）

    Args:
        sheet: Excel 工作表物件
        current_row: 當前列號
        current_col: 當前欄號

    Returns:
        (new_row, new_col): 新的列號和欄號
    """
    line_no = get_line_number(current_row)
    if line_no > 1:
        # 移動到上一行的相同欄位
        new_line = line_no - 1
        new_row = get_row_from_line(new_line)
        new_col = current_col

        # 檢查目標儲存格是否有效（不超過行尾）
        if new_col > END_COL:
            new_col = END_COL

        return new_row, new_col
    else:
        # 已在第一行，不移動
        return current_row, current_col


def move_down(sheet, current_row: int, current_col: int, total_lines: int) -> tuple:
    """
    向下移動游標到下一行的相同欄位（或行首）

    Args:
        sheet: Excel 工作表物件
        current_row: 當前列號
        current_col: 當前欄號
        total_lines: 總行數

    Returns:
        (new_row, new_col): 新的列號和欄號
    """
    line_no = get_line_number(current_row)
    if line_no < total_lines:
        # 移動到下一行的相同欄位
        new_line = line_no + 1
        new_row = get_row_from_line(new_line)
        new_col = current_col

        # 檢查目標儲存格是否有效（不超過行尾）
        if new_col > END_COL:
            new_col = END_COL

        return new_row, new_col
    else:
        # 已在最後一行，不移動
        return current_row, current_col


def move_left(sheet, current_row: int, current_col: int) -> tuple:
    """
    向左移動游標

    Args:
        sheet: Excel 工作表物件
        current_row: 當前列號
        current_col: 當前欄號

    Returns:
        (new_row, new_col): 新的列號和欄號
    """
    # 如果已在行首，則跳到前一行的行尾
    if current_col == START_COL:
        line_no = get_line_number(current_row)
        if line_no > 1:
            # 跳到前一行，找到最後一個有內容的儲存格
            new_line = line_no - 1
            new_row = get_row_from_line(new_line)

            # 從行尾往回找，找到第一個有內容或換行符的儲存格
            for col in range(END_COL, START_COL - 1, -1):
                cell = sheet.range((new_row, col))
                cell_value = cell.value
                cell_formula = cell.formula

                # 如果是換行符，跳過
                if cell_formula and "=CHAR(10)" in cell_formula.upper():
                    continue
                if cell_value == "\n":
                    continue

                # 找到有內容的儲存格
                if cell_value is not None and str(cell_value).strip():
                    return new_row, col

            # 如果都沒有內容，就跳到行首
            return new_row, START_COL
        else:
            # 已在第一行行首，不移動
            return current_row, current_col
    else:
        # 在行中，向左移動一格
        return current_row, current_col - 1


def move_right(sheet, current_row: int, current_col: int, total_lines: int) -> tuple:
    """
    向右移動游標

    注意：若下一個儲存格為換行控制碼（\\n），則跳到下一行行首

    Args:
        sheet: Excel 工作表物件
        current_row: 當前列號
        current_col: 當前欄號
        total_lines: 總行數

    Returns:
        (new_row, new_col): 新的列號和欄號
    """
    # 先檢查是否已在行尾
    if current_col >= END_COL:
        line_no = get_line_number(current_row)
        if line_no < total_lines:
            # 跳到下一行的行首
            new_line = line_no + 1
            new_row = get_row_from_line(new_line)
            new_col = START_COL
            # print(f"  [已到行尾 {xw.utils.col_name(current_col)}{current_row}，跳到下一行 {xw.utils.col_name(new_col)}{new_row}]")
            return new_row, new_col
        else:
            # 已在最後一行行尾，不移動
            # print(f"  [已在最後一行行尾，無法繼續向右]")
            return current_row, current_col

    # 檢查下一格
    next_col = current_col + 1
    next_cell = sheet.range((current_row, next_col))
    next_cell_value = next_cell.value
    next_cell_formula = next_cell.formula

    # 調試輸出
    # print(f"  [檢查下一格 {xw.utils.col_name(next_col)}{current_row}]")
    # print(f"    值: {repr(next_cell_value)}")
    # print(f"    公式: {next_cell_formula}")

    # 檢查是否為換行控制碼
    is_newline = False

    # 方法1: 檢查公式是否為 =CHAR(10)
    if next_cell_formula and "=CHAR(10)" in next_cell_formula.upper():
        is_newline = True
        # print(f"    → 偵測到 CHAR(10) 公式")

    # 方法2: 檢查值是否為換行符
    elif next_cell_value is not None:
        if next_cell_value == "\n" or next_cell_value == chr(10):
            is_newline = True
            # print(f"    → 偵測到換行符值")

    if is_newline:
        # 遇到換行符，跳到下一行行首
        line_no = get_line_number(current_row)
        if line_no < total_lines:
            new_line = line_no + 1
            new_row = get_row_from_line(new_line)
            new_col = START_COL
            # print(f"  [偵測到換行符，跳到下一行 {xw.utils.col_name(new_col)}{new_row}]")
            return new_row, new_col
        else:
            # 已在最後一行，不移動
            # print(f"  [已在最後一行，無法跳到下一行]")
            return current_row, current_col
    else:
        # 正常向右移動一格
        # print(f"  [正常向右移動到 {xw.utils.col_name(next_col)}{current_row}]")
        return current_row, next_col


def get_total_lines(wb) -> int:
    """
    取得總行數

    Args:
        wb: Excel 工作簿物件

    Returns:
        總行數
    """
    try:
        total_lines = int(wb.names["每頁總列數"].refers_to_range.value)
        return total_lines
    except:  # noqa: E722
        # 預設值
        return 10


def get_lines_per_page(wb) -> int:
    """
    取得【每頁行數】：以 Excel 視窗目前可視列數，換算可容納幾【行】
    （每行佔 ROWS_PER_LINE 列）；供 PgUp/PgDn 翻頁使用。

    Args:
        wb: Excel 工作簿物件

    Returns:
        每頁行數（至少 1 行；無法取得時，預設 5 行）
    """
    try:
        visible_rows = wb.app.api.ActiveWindow.VisibleRange.Rows.Count
        return max(1, visible_rows // ROWS_PER_LINE)
    except Exception as e:
        logging.debug(f"無法取得可視列數：{e}")
        return 5


def hide_manual_annotation_style(wb):
    """
    隱藏【人工標音儲存格】樣式的文字
    將字型顏色改為與填滿顏色相同（象牙白）

    Args:
        wb: Excel 工作簿物件
    """
    try:
        # 取得 Excel API 物件
        # excel_app = wb.app.api
        workbook = wb.api

        # 查找【人工標音儲存格】樣式
        style_name = "人工標音儲存格"
        try:
            style = workbook.Styles(style_name)
            # 將字型顏色改為象牙白（RGB: 255, 255, 240）
            # Excel 使用 BGR 格式，所以順序相反
            style.Font.Color = 0xF0FFFF  # BGR: 240, 255, 255 (象牙白)
            print(f"✓ 已隱藏【{style_name}】樣式的文字（字型顏色改為象牙白）")
        except Exception as e:
            print(f"⚠️  找不到【{style_name}】樣式，跳過隱藏操作：{e}！")

    except Exception as e:
        logging.warning(f"隱藏人工標音樣式失敗：{e}")
        print(f"⚠️  隱藏人工標音樣式失敗：{e}")


def restore_manual_annotation_style(wb):
    """
    恢復【人工標音儲存格】樣式的文字
    將字型顏色改回紅色

    Args:
        wb: Excel 工作簿物件
    """
    try:
        # 取得 Excel API 物件
        # excel_app = wb.app.api
        workbook = wb.api

        # 查找【人工標音儲存格】樣式
        style_name = "人工標音儲存格"
        try:
            style = workbook.Styles(style_name)
            # 將字型顏色改回紅色（RGB: 255, 0, 0）
            # Excel 使用 BGR 格式，所以順序相反
            style.Font.Color = 0x0000FF  # BGR: 0, 0, 255 (紅色)
            print(f"✓ 已恢復【{style_name}】樣式的文字（字型顏色改回紅色）")
        except Exception as e:
            print(f"⚠️  找不到【{style_name}】樣式，跳過恢復操作：{e}！")

    except Exception as e:
        logging.warning(f"恢復人工標音樣式失敗：{e}")
        print(f"⚠️  恢復人工標音樣式失敗：{e}")


# =========================================================================
# 視窗切換函數
# =========================================================================
def activate_excel_window(wb):
    """
    激活 Excel 視窗，使其成為前景視窗

    Args:
        wb: Excel 工作簿物件
    """
    if not HAS_WIN32:
        print("提示：無法自動切換到 Excel 視窗（需要 pywin32 套件）")
        print("請手動點擊 Excel 視窗以顯示十字游標")
        return

    try:
        # 取得 Excel 視窗句柄
        excel_hwnd = wb.app.api.Hwnd

        # 檢查視窗是否存在
        if not win32gui.IsWindow(excel_hwnd):
            print("無法找到 Excel 視窗")
            return

        # 如果視窗最小化，先還原
        if win32gui.IsIconic(excel_hwnd):
            win32gui.ShowWindow(excel_hwnd, win32con.SW_RESTORE)

        # 將 Excel 視窗切換到前景
        win32gui.SetForegroundWindow(excel_hwnd)
        print("✓ 已切換到 Excel 視窗")

        # 等待視窗切換完成
        time.sleep(0.5)

    except Exception as e:
        logging.error(f"無法激活 Excel 視窗：{e}")


# =========================================================================
# 主要處理函數（使用鍵盤監聽）
# =========================================================================
class NavigationController:
    """導航控制器 - 使用鍵盤監聽"""

    def __init__(self, wb, sheet, edit_mode=False, console_hwnd=None):
        self.wb = wb
        self.sheet = sheet
        self.edit_mode = edit_mode  # 是否為校稿模式
        self.current_row = START_ROW
        self.current_col = START_COL
        self.total_lines = get_total_lines(wb)
        self.running = True
        self.pending_action = None  # 待執行的動作
        self.listener = None  # 鍵盤監聽器
        self.last_move_time = None  # 上次移動時間（用於延遲檢查）
        self.auto_skip_delay = 0.5  # 自動跳過換行的延遲時間（秒）
        self.auto_skip_enabled = True  # 是否啟用自動跳過換行

        # 儲存視窗句柄（用於切換視窗）
        self.console_hwnd = console_hwnd
        self.excel_hwnd = None
        if HAS_WIN32:
            try:
                self.excel_hwnd = wb.app.api.Hwnd
                if not self.console_hwnd:
                    # 若啟動時已切到 Excel，改依 WezTerm／Terminal 行程補抓
                    self.console_hwnd = capture_console_hwnd(excel_hwnd=self.excel_hwnd)
                logging.info(f"Console 視窗句柄：{self.console_hwnd}")
                logging.info(f"Excel 視窗句柄：{self.excel_hwnd}")
            except Exception as e:
                logging.warning(f"無法取得視窗句柄：{e}")

    def move_to_cell(self, row, col, reset_timer=True):
        """移動到指定儲存格"""
        self.current_row = row
        self.current_col = col
        self.sheet.range((row, col)).select()

        # 記錄移動時間，用於延遲檢查
        if reset_timer:
            self.last_move_time = time.time()

        # 顯示當前位置
        current_cell = self.sheet.range((row, col))
        cell_value = current_cell.value
        line_no = get_line_number(row)
        col_letter = xw.utils.col_name(col)
        display_value = cell_value or ""
        print(f"→ 第 {line_no} 行，{col_letter}{row}【{display_value}】")

    def current_cell_address(self) -> str:
        """目前導航位置的 Excel 位址，例如 F5。"""
        return f"{xw.utils.col_name(self.current_col)}{self.current_row}"

    def build_query_args(self, manual_input: bool = False):
        """建立傳給 a250／a260 的參數，明確帶入目前漢字儲存格與終端機視窗。"""
        return argparse.Namespace(
            new=False,
            cell=self.current_cell_address(),
            console_hwnd=self.console_hwnd,
            manual_input=manual_input,
        )

    def select_current_cell(self):
        """將 Excel 作用儲存格對齊目前導航位置。"""
        try:
            self.sheet.activate()
            self.sheet.range((self.current_row, self.current_col)).select()
        except Exception as e:
            logging.debug(f"重新選取目前儲存格失敗：{e}")

    def focus_console(self, quiet: bool = False):
        """把焦點切回啟動本程式的終端機（含 WezTerm）。"""
        activate_console_window(self.console_hwnd, excel_hwnd=self.excel_hwnd, quiet=quiet)

    def check_and_skip_newline(self):
        """檢查當前儲存格是否為換行符號，如果是則自動跳到下一行"""
        if not self.auto_skip_enabled:
            return

        # 檢查是否已經過了延遲時間
        if self.last_move_time is None:
            return

        elapsed = time.time() - self.last_move_time
        if elapsed < self.auto_skip_delay:
            return  # 還沒到延遲時間

        # 延遲時間已到，檢查當前儲存格
        try:
            current_cell = self.sheet.range((self.current_row, self.current_col))
            cell_value = current_cell.value
            cell_formula = current_cell.formula
        except Exception as e:
            # 查字過程若短暫中斷 Excel COM，不應讓整支導航程式結束
            logging.debug(f"檢查換行符號失敗：{e}")
            self.last_move_time = None
            return

        is_newline = False
        # 檢查公式是否為 =CHAR(10)
        if cell_formula and "=CHAR(10)" in str(cell_formula).upper():
            is_newline = True
        # 檢查值是否為換行符
        elif cell_value is not None:
            if cell_value == "\n" or cell_value == chr(10):
                is_newline = True

        if is_newline:
            # 當前儲存格是換行符號，自動跳到下一行
            line_no = get_line_number(self.current_row)
            if line_no < self.total_lines:
                print("  [偵測到換行符號，自動跳到下一行]")
                new_line = line_no + 1
                new_row = get_row_from_line(new_line)
                new_col = START_COL
                # 移動到下一行，不重置計時器避免無限循環
                self.move_to_cell(new_row, new_col, reset_timer=False)
                # 清除計時器
                self.last_move_time = None

    def on_key_press(self, key):
        """鍵盤按下事件處理 - 只設置動作標記"""
        try:
            if key == keyboard.Key.left:
                self.pending_action = "left"
            elif key == keyboard.Key.right:
                self.pending_action = "right"
            elif key == keyboard.Key.up:
                self.pending_action = "up"
            elif key == keyboard.Key.down:
                self.pending_action = "down"
            elif key == keyboard.Key.page_up:
                # PgUp 鍵：翻到上一頁
                self.pending_action = "page_up"
            elif key == keyboard.Key.page_down:
                # PgDn 鍵：翻到下一頁
                self.pending_action = "page_down"
            elif key == keyboard.Key.space:
                # 空白鍵：查詢個人字典
                # self.pending_action = "query_personal"
                self.pending_action = "query_and_replace_thok_im"
            elif hasattr(key, "char") and key.char:
                # 處理字元鍵
                if key.char.lower() == " ":
                    # 空白鍵：查詢個人字典
                    self.pending_action = "query_and_replace_thok_im"
                elif key.char.lower() == "j":
                    # J 鍵：更換漢字讀音
                    self.pending_action = "query_and_assign_thok_im"
                # elif key.char.lower() == "s":
                #     # S 鍵：查詢萌典
                #     self.pending_action = "query_moedict"
                elif key.char.lower() == "e":
                    # E 鍵：手動輸入人工標音
                    self.pending_action = "manual_input"
                elif key.char == "=":
                    # = 鍵：引用既有的人工標音
                    self.pending_action = "fill_manual_mark"
            elif key == keyboard.Key.delete:
                # Del 鍵：清除人工標音
                self.pending_action = "clear_manual_annotation"
            elif key == keyboard.Key.esc:
                self.pending_action = "esc"
                self.running = False
                return False  # 停止監聽
        except AttributeError:
            pass
        except Exception as e:
            logging.error(f"按鍵處理錯誤：{e}")

    def process_pending_action(self):
        """處理待執行的動作(在主執行緒中執行)"""
        if self.pending_action is None:
            return

        action = self.pending_action
        self.pending_action = None  # 清除動作

        try:
            if action == "left":
                # 向左移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_left(self.sheet, self.current_row, self.current_col)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == "right":
                # 向右移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_right(self.sheet, self.current_row, self.current_col, self.total_lines)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == "up":
                # 向上移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_up(self.sheet, self.current_row, self.current_col)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == "down":
                # 向下移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_down(self.sheet, self.current_row, self.current_col, self.total_lines)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == "page_up":
                # 翻到上一頁（重置延遲計時器）
                self.last_move_time = None
                lines_per_page = get_lines_per_page(self.wb)
                line_no = get_line_number(self.current_row)
                new_line = max(1, line_no - lines_per_page)
                if new_line != line_no:
                    print(f"  [PgUp：向上翻 {line_no - new_line} 行]")
                    self.move_to_cell(get_row_from_line(new_line), self.current_col)
                else:
                    print("  [已在第一行，無法向上翻頁]")

            elif action == "page_down":
                # 翻到下一頁（重置延遲計時器）
                self.last_move_time = None
                lines_per_page = get_lines_per_page(self.wb)
                line_no = get_line_number(self.current_row)
                new_line = min(self.total_lines, line_no + lines_per_page)
                if new_line != line_no:
                    print(f"  [PgDn：向下翻 {new_line - line_no} 行]")
                    self.move_to_cell(get_row_from_line(new_line), self.current_col)
                else:
                    print("  [已在最後一行，無法向下翻頁]")

            # elif action == "query_moedict":
            #     # 查詢萌典
            #     self.query_moedict_dictionary()
            #
            # elif action == "query_and_replace_thok_im":
            #     # 查詢個人字典
            #     self.query_personal_dictionary()
            #
            elif action == "query_and_replace_thok_im":
                # 更換漢字讀音
                self.query_dictionary_and_replace_han_ji_thok_im()

            elif action == "query_and_assign_thok_im":
                # 指定漢字讀音
                self.query_dictionary_and_assign_han_ji_thok_im()

            elif action == "fill_manual_mark":
                # 填入人工標音標記
                self.fill_manual_annotation_mark()

            elif action == "manual_input":
                # 手動輸入人工標音
                self.manual_input_annotation()

            elif action == "clear_manual_annotation":
                # 清除人工標音
                self.clear_manual_annotation()

            elif action == "esc":
                print("\n按下 ESC 鍵，程式結束")

        except Exception as e:
            logging.error(f"執行動作錯誤：{e}")

    # def query_moedict_dictionary(self):
    #     """查詢萌典字典"""
    #     print("\n" + "=" * 70)
    #     print("進入萌典字典查詢模式")
    #     print("=" * 70)
    #
    #     # 暫停鍵盤監聽
    #     if self.listener:
    #         self.listener.stop()
    #         time.sleep(0.3)
    #
    #     try:
    #         if HAS_A220:
    #             # 直接調用 a220 的核心函數，不進入無限循環
    #             print("\n查詢萌典字典中...")
    #
    #             # 切換到終端機視窗（確保用戶可以輸入）
    #             activate_console_window(self.console_hwnd)
    #
    #             # # 取得設定值
    #             # try:
    #             #     from mod_excel_access import get_value_by_name
    #
    #             #     ue_im_lui_piat = get_value_by_name(wb=self.wb, name="語音類型")
    #             #     han_ji_khoo = get_value_by_name(wb=self.wb, name="漢字庫")
    #             # except:
    #             #     ue_im_lui_piat = "白話音"
    #             #     han_ji_khoo = "河洛話"
    #
    #             # 取得當前作用儲存格位置
    #             current_cell = f"{xw.utils.col_name(self.current_col)}{self.current_row}"
    #             print(f"當前儲存格：{current_cell}")
    #
    #             # 調用查詢函數
    #             # exit_code = ca_han_ji_thak_im_a220(
    #             #     wb=self.wb,
    #             #     sheet_name='漢字注音',
    #             #     cell=current_cell,
    #             #     ue_im_lui_piat=ue_im_lui_piat,
    #             #     han_ji_khoo=han_ji_khoo,
    #             #     new_khuat_ji_piau_sheet=False,
    #             #     new_piau_im_ji_khoo_sheet=False,
    #             # )
    #             exit_code = ca_han_ji_thak_im_a220(
    #                 wb=self.wb,
    #                 args=None,
    #             )
    #
    #             if exit_code == 0:
    #                 print("\n✓ 查詢完成")
    #             else:
    #                 print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
    #         else:
    #             # 回退到 subprocess 方式
    #             print("\n執行 a220_作用儲存格查找萌典漢字讀音.py...")
    #             result = subprocess.run(
    #                 [sys.executable, "a220_作用儲存格查找萌典漢字讀音.py"],
    #                 cwd=os.path.dirname(os.path.abspath(__file__)),
    #                 capture_output=False,
    #                 text=True,
    #             )
    #             if result.returncode != 0:
    #                 print(f"⚠️  a220 程式執行失敗，返回碼：{result.returncode}")
    #     except KeyboardInterrupt:
    #         print("\n\n使用者中斷查詢")
    #     except Exception as e:
    #         logging.error(f"執行萌典查詢失敗：{e}")
    #         print(f"❌ 執行萌典查詢失敗：{e}")
    #     finally:
    #         print("\n" + "=" * 70)
    #         print("返回導航模式")
    #         print("=" * 70)
    #
    #         # 切換回 Excel 視窗
    #         activate_excel_window(self.wb)
    #
    #         # 重新啟動鍵盤監聽
    #         if self.listener:
    #             self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
    #             self.listener.start()
    #             time.sleep(0.3)
    #         print("✓ 已恢復導航模式\n")

    def query_dictionary_and_replace_han_ji_thok_im(self):
        """查詢字典將漢字讀音更換"""
        print("\n" + "=" * 70)
        print("進入【漢字更換讀音】模式")
        print("=" * 70)

        # 暫停鍵盤監聽
        if self.listener:
            self.listener.stop()
            time.sleep(0.3)

        try:
            if HAS_A250:
                # 直接調用 a250 的核心函數，不進入無限循環
                print("\n查詢字典中...")

                # 先對齊 Excel 作用儲存格，再切換到終端機視窗（確保用戶可以輸入）
                self.select_current_cell()
                self.focus_console()

                # 取得當前作用儲存格位置
                current_cell = self.current_cell_address()
                print(f"當前儲存格：{current_cell}")

                # 調用查詢函數（明確傳入目前漢字儲存格，避免初始化後讀到錯誤位置）
                exit_code = ca_han_ji_thok_im_a250(
                    wb=self.wb,
                    args=self.build_query_args(),
                )

                if exit_code == 0:
                    print("\n✓ 查詢完成")
                else:
                    print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
            else:
                # 回退到 subprocess 方式
                print("\n執行 a250_更換漢字讀音.py...")
                result = subprocess.run(
                    [sys.executable, "a250_更換漢字讀音.py"],
                    cwd=os.path.dirname(os.path.abspath(__file__)),
                    capture_output=False,
                    text=True,
                )
                if result.returncode != 0:
                    print(f"⚠️  a250 程式執行失敗，返回碼：{result.returncode}")
        except KeyboardInterrupt:
            print("\n\n使用者中斷查詢")
        except Exception as e:
            logging.error(f"執行字典查詢失敗：{e}")
            print(f"❌ 執行字典查詢失敗：{e}")
        finally:
            print("\n" + "=" * 70)
            print("返回導航模式")
            print("=" * 70)

            # 切換回 Excel 視窗，並回到查字前的漢字儲存格
            self.select_current_cell()
            activate_excel_window(self.wb)

            # 重新啟動鍵盤監聽
            if self.listener:
                self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
                self.listener.start()
                time.sleep(0.3)
            print("✓ 已恢復導航模式\n")

    def query_dictionary_and_assign_han_ji_thok_im(self):
        """查詢字典指定漢字讀音"""
        print("\n" + "=" * 70)
        print("進入【漢字指定讀音】模式")
        print("=" * 70)

        # 暫停鍵盤監聽
        if self.listener:
            self.listener.stop()
            time.sleep(0.3)

        try:
            if HAS_A260:
                # 直接調用 a260 的核心函數，不進入無限循環
                print("\n查詢字典中...")

                # 先對齊 Excel 作用儲存格，再切換到終端機視窗（確保用戶可以輸入）
                self.select_current_cell()
                self.focus_console()

                # 取得當前作用儲存格位置
                current_cell = self.current_cell_address()
                print(f"當前儲存格：{current_cell}")

                # 調用查詢函數（明確傳入目前漢字儲存格，避免初始化後讀到錯誤位置）
                exit_code = ca_ji_tian_au_thiam_jin_kang_piau_im(
                    wb=self.wb,
                    args=self.build_query_args(),
                )

                if exit_code == 0:
                    print("\n✓ 查詢完成")
                else:
                    print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
            else:
                # 回退到 subprocess 方式
                print("\n執行 a260_為單一漢字指定讀音.py...")
                result = subprocess.run(
                    [sys.executable, "a260_為單一漢字指定讀音.py"],
                    cwd=os.path.dirname(os.path.abspath(__file__)),
                    capture_output=False,
                    text=True,
                )
                if result.returncode != 0:
                    print(f"⚠️  a222 程式執行失敗，返回碼：{result.returncode}")
        except KeyboardInterrupt:
            print("\n\n使用者中斷查詢")
        except Exception as e:
            logging.error(f"執行字典查詢失敗：{e}")
            print(f"❌ 執行字典查詢失敗：{e}")
        finally:
            print("\n" + "=" * 70)
            print("返回導航模式")
            print("=" * 70)

            # 切換回 Excel 視窗，並回到查字前的漢字儲存格
            self.select_current_cell()
            activate_excel_window(self.wb)

            # 重新啟動鍵盤監聽
            if self.listener:
                self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
                self.listener.start()
                time.sleep(0.3)
            print("✓ 已恢復導航模式\n")

    # def query_personal_dictionary(self):
    #     """查詢個人字典"""
    #     print("\n" + "=" * 70)
    #     print("進入個人字典查詢模式")
    #     print("=" * 70)
    #
    #     # 暫停鍵盤監聽
    #     if self.listener:
    #         self.listener.stop()
    #         time.sleep(0.3)
    #
    #     try:
    #         if HAS_A222:
    #             # 直接調用 a222 的核心函數，不進入無限循環
    #             print("\n查詢個人字典中...")
    #
    #             # 切換到終端機視窗（確保用戶可以輸入）
    #             activate_console_window(self.console_hwnd)
    #
    #             # 取得當前作用儲存格位置
    #             current_cell = f"{xw.utils.col_name(self.current_col)}{self.current_row}"
    #             print(f"當前儲存格：{current_cell}")
    #
    #             # 調用查詢函數
    #             exit_code = ca_han_ji_thak_im_a222(
    #                 wb=self.wb,
    #                 args=None,
    #             )
    #
    #             if exit_code == 0:
    #                 print("\n✓ 查詢完成")
    #             else:
    #                 print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
    #         else:
    #             # 回退到 subprocess 方式
    #             print("\n執行 a222_依作用儲存格在個人字典查找漢字讀音.py...")
    #             result = subprocess.run(
    #                 [sys.executable, "a222_依作用儲存格在個人字典查找漢字讀音.py"],
    #                 cwd=os.path.dirname(os.path.abspath(__file__)),
    #                 capture_output=False,
    #                 text=True,
    #             )
    #             if result.returncode != 0:
    #                 print(f"⚠️  a222 程式執行失敗，返回碼：{result.returncode}")
    #     except KeyboardInterrupt:
    #         print("\n\n使用者中斷查詢")
    #     except Exception as e:
    #         logging.error(f"執行個人字典查詢失敗：{e}")
    #         print(f"❌ 執行個人字典查詢失敗：{e}")
    #     finally:
    #         print("\n" + "=" * 70)
    #         print("返回導航模式")
    #         print("=" * 70)
    #
    #         # 切換回 Excel 視窗
    #         activate_excel_window(self.wb)
    #
    #         # 重新啟動鍵盤監聽
    #         if self.listener:
    #             self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
    #             self.listener.start()
    #             time.sleep(0.3)
    #         print("✓ 已恢復導航模式\n")

    def fill_manual_annotation_mark(self):
        """填入人工標音標記【=】到當前儲存格上方兩列的人工標音儲存格，並執行 a224 查詢更新標音"""
        try:
            # 計算人工標音儲存格的位置（當前儲存格上方兩列）
            manual_annotation_row = self.current_row - 2
            manual_annotation_col = self.current_col

            # 確認位置有效
            if manual_annotation_row < 1:
                print("⚠️  無法填入：當前位置沒有人工標音儲存格")
                return

            # 填入【=】字元
            target_cell = self.sheet.range((manual_annotation_row, manual_annotation_col))
            target_cell.value = "="

            # 顯示訊息
            col_letter = xw.utils.col_name(manual_annotation_col)
            current_cell_address = f"{xw.utils.col_name(self.current_col)}{self.current_row}"
            target_cell_address = f"{col_letter}{manual_annotation_row}"

            print(f"\n✓ 已在 {target_cell_address} 填入人工標音標記【=】")
            print(f"  (當前漢字儲存格：{current_cell_address})")

            # 執行 a224 查詢以更新標音
            print("\n" + "=" * 70)
            print("執行 a224 程式：引用既有的漢字標音")
            print("=" * 70)

            # 暫停鍵盤監聽
            if self.listener:
                self.listener.stop()
                time.sleep(0.3)

            try:
                if HAS_A224:
                    # 直接調用 a224 的核心函數
                    print("\n查詢並更新標音中...")

                    # 切換到終端機視窗（確保用戶可以輸入）
                    self.focus_console()

                    # 取得設定值
                    # try:
                    #     from mod_excel_access import get_value_by_name
                    #     ue_im_lui_piat = get_value_by_name(wb=self.wb, name='語音類型')
                    #     han_ji_khoo = get_value_by_name(wb=self.wb, name='漢字庫')
                    # except:
                    #     ue_im_lui_piat = "白話音"
                    #     han_ji_khoo = "河洛話"

                    # 取得當前作用儲存格位置
                    current_cell = self.current_cell_address()
                    print(f"當前儲存格：{current_cell}")

                    # 調用查詢函數
                    exit_code = jin_kang_piau_im_ca_taigi_im_piau(
                        wb=self.wb,
                        args=self.build_query_args(),
                    )

                    if exit_code == 0:
                        print("\n✓ 查詢完成")
                    else:
                        print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
                else:
                    # 回退到 subprocess 方式
                    print("\n執行 a224_引用既有的漢字標音.py...")
                    result = subprocess.run(
                        [sys.executable, "a224_引用既有的漢字標音.py"],
                        cwd=os.path.dirname(os.path.abspath(__file__)),
                        capture_output=False,
                        text=True,
                    )
                    if result.returncode != 0:
                        print(f"⚠️  a224 程式執行失敗，返回碼：{result.returncode}")
            except KeyboardInterrupt:
                print("\n\n使用者中斷查詢")
            except Exception as e:
                logging.error(f"執行 a224 查詢失敗：{e}")
                print(f"❌ 執行 a224 查詢失敗：{e}")
            finally:
                print("\n" + "=" * 70)
                print("返回導航模式")
                print("=" * 70)

                # 切換回 Excel 視窗
                activate_excel_window(self.wb)

                if self.listener:
                    self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
                    self.listener.start()
                    time.sleep(0.3)
                print("✓ 已恢復導航模式\n")

        except Exception as e:
            logging.error(f"填入人工標音標記失敗：{e}")
            print(f"\n❌ 填入失敗：{e}\n")
            # 確保恢復鍵盤監聽
            if self.listener:
                self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
                self.listener.start()

    def manual_input_annotation(self):
        """手動輸入人工標音到當前儲存格上方兩列的人工標音儲存格"""
        try:
            # 計算人工標音儲存格的位置（當前儲存格上方兩列）
            manual_annotation_row = self.current_row - 2
            manual_annotation_col = self.current_col

            # 確認位置有效
            if manual_annotation_row < 1:
                print("⚠️  無法輸入：當前位置沒有人工標音儲存格")
                return

            # 顯示當前儲存格資訊
            col_letter = xw.utils.col_name(manual_annotation_col)
            current_cell_address = f"{xw.utils.col_name(self.current_col)}{self.current_row}"
            target_cell_address = f"{col_letter}{manual_annotation_row}"
            current_han_ji = self.sheet.range((self.current_row, self.current_col)).value or ""

            print("\n" + "=" * 70)
            print("手動輸入人工標音模式")
            print("=" * 70)
            print(f"當前漢字儲存格：{current_cell_address}【{current_han_ji}】")
            print(f"人工標音儲存格：{target_cell_address}")
            print("\n輸入說明：")
            print("  - 可輸入帶調符的台羅拼音（如：Tông, Sióng）")
            print("  - 可輸入帶調號的台羅拼音（如：Tong5, Siong2）")
            print("  - 可使用 Ctrl+V 貼上複製的內容")
            print("  - 按 Enter 確認輸入，直接按 Enter 則放棄輸入")
            print("=" * 70)

            # 暫停鍵盤監聽
            if self.listener:
                self.listener.stop()
                time.sleep(0.3)

            # 確保切換回終端機，以便使用者輸入文字
            self.focus_console()

            try:
                if HAS_A260:
                    # E 鍵：略過查字典，直接手動輸入人工標音
                    exit_code = ca_ji_tian_au_thiam_jin_kang_piau_im(
                        wb=self.wb,
                        args=self.build_query_args(manual_input=True),
                    )
                    if exit_code == 0:
                        print("✓ 已完成台語音標與漢字標音更新")
                    else:
                        print(f"⚠️  更新結果：exit_code = {exit_code}")
                else:
                    print("⚠️  a260 模組未載入，無法更新")

            except KeyboardInterrupt:
                print("\n\n使用者中斷輸入")
            except Exception as e:
                logging.error(f"手動輸入人工標音失敗：{e}")
                print(f"\n❌ 輸入失敗：{e}")
            finally:
                print("\n" + "=" * 70)
                print("返回導航模式")
                print("=" * 70)

                # 切換回 Excel 視窗
                activate_excel_window(self.wb)

                # 重新啟動鍵盤監聽
                if self.listener:
                    self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
                    self.listener.start()
                    time.sleep(0.3)
                print("✓ 已恢復導航模式\n")

        except Exception as e:
            logging.error(f"手動輸入人工標音失敗：{e}")
            print(f"\n❌ 輸入失敗：{e}\n")
            # 確保恢復鍵盤監聽
            if self.listener:
                self.listener = keyboard.Listener(on_press=self.on_key_press, suppress=True)
                self.listener.start()

    def clear_manual_annotation(self):
        """清除當前儲存格上方兩列的人工標音儲存格內容"""
        try:
            # 計算人工標音儲存格的位置（當前儲存格上方兩列）
            manual_annotation_row = self.current_row - 2
            manual_annotation_col = self.current_col

            # 確認位置有效
            if manual_annotation_row < 1:
                print("\n⚠️  無法清除：當前位置沒有人工標音儲存格")
                return

            # 取得儲存格資訊
            col_letter = xw.utils.col_name(manual_annotation_col)
            target_cell_address = f"{col_letter}{manual_annotation_row}"
            target_cell = self.sheet.range((manual_annotation_row, manual_annotation_col))
            current_value = target_cell.value or ""

            # 如果儲存格已經是空的
            if not current_value:
                print(f"\n⚠️  人工標音儲存格 {target_cell_address} 已經是空的")
                return

            print(f"\n清除人工標音：{target_cell_address}【{current_value}】")

            # 清除儲存格內容
            # target_cell.value = ""
            target_cell.value = "#"
            print(f"✓ 已清除 {target_cell_address} 的人工標音")

            # 呼叫 a224 程式以更新台語音標與漢字標音
            print("\n正在更新台語音標與漢字標音...")
            try:
                if HAS_A224:
                    # 創建簡單的 args 物件（模擬命令列參數），並帶入目前漢字儲存格
                    exit_code = jin_kang_piau_im_ca_taigi_im_piau(wb=self.wb, args=self.build_query_args())
                    if exit_code == 0:
                        print("✓ 已完成台語音標與漢字標音更新\n")
                    else:
                        print(f"⚠️  更新結果：exit_code = {exit_code}\n")
                else:
                    print("⚠️  a224 模組未載入，無法更新\n")
            except Exception as e:
                logging.error(f"更新台語音標與漢字標音失敗：{e}")
                print(f"❌ 更新失敗：{e}\n")

        except Exception as e:
            logging.error(f"清除人工標音失敗：{e}")
            print(f"\n❌ 清除失敗：{e}\n")


def read_han_ji_with_keyboard(wb, view_mode=False, console_hwnd=None, start_cell=None) -> int:
    """
    漢字注音工作表導讀主程式（使用鍵盤監聽）

    Args:
        wb: Excel 工作簿物件
        view_mode: 是否為瀏覽模式（True=隱藏人工標音；False=校對模式）
        console_hwnd: 啟動本程式的終端機視窗句柄（WezTerm 等）
        start_cell: 校對起始儲存格（如 d133）；未指定時自 D5 開始

    Returns:
        退出代碼
    """
    try:
        # 取得工作表
        sheet = wb.sheets[SHEET_NAME]
        sheet.activate()

        # 初始化控制器
        controller = NavigationController(wb, sheet, edit_mode=view_mode, console_hwnd=console_hwnd)

        try:
            start_row, start_col = resolve_start_cell(start_cell, controller.total_lines)
        except ValueError as e:
            print(f"❌ {e}")
            return EXIT_CODE_INVALID_INPUT

        # 移動到指定起始儲存格（預設第一行行首 D5）
        controller.move_to_cell(start_row, start_col)
        if start_cell:
            requested = str(start_cell).strip().upper().replace("$", "")
            actual = f"{xw.utils.col_name(start_col)}{start_row}"
            if requested != actual:
                print(f"起始儲存格已對齊漢字列：{requested} → {actual}")
            else:
                print(f"起始儲存格：{actual}")

        print("=" * 70)
        if view_mode:
            print("漢字注音工作表導讀（鍵盤監聽模式 - 校稿模式）")
        else:
            print("漢字注音工作表導讀（鍵盤監聽模式）")
        print("=" * 70)
        print("操作說明：")
        print("  ← (Left Arrow)  : 向左移動")
        print("  → (Right Arrow) : 向右移動")
        print("  ↑ (Up Arrow)    : 向上移動到上一行")
        print("  ↓ (Down Arrow)  : 向下移動到下一行")
        print("  PgUp (Page Up)  : 翻到上一頁")
        print("  PgDn (Page Down): 翻到下一頁")
        print("  空白 鍵         : 查字典更換漢字讀音（a250）")
        print("  J 鍵            : 查字典指定漢字讀音（a260）")
        print("  E 鍵            : 手動輸入人工標音（a222）")
        print("  = 鍵            : 填入人工標音標記")
        print("  ESC             : 結束程式")
        print("=" * 70)
        print(f"總行數：{controller.total_lines}")
        print(f"每行字數：{END_COL - START_COL + 1}")
        if view_mode:
            print("模式：瀏覽模式（隱藏人工標音顏色）")
        else:
            print("模式：校對模式（顯示人工標音）")
        print("=" * 70)

        # 【瀏覽模式】需隱藏【人工標音】
        if view_mode:
            print("\n正在隱藏人工標音文字...")
            hide_manual_annotation_style(wb)

        # 切換到 Excel 視窗，讓十字游標顯示
        print("\n正在切換到 Excel 視窗...")
        activate_excel_window(wb)

        print("\n請使用方向鍵導航...")
        print("提示：程式會攔截按鍵，不會影響 Excel 儲存格內容")

        # 啟動鍵盤監聽（在背景執行緒，使用 suppress=True 攔截所有按鍵）
        controller.listener = keyboard.Listener(
            on_press=controller.on_key_press,
            suppress=True,  # 攔截按鍵，不讓 Excel 接收
        )
        controller.listener.start()

        try:
            # 主迴圈：在主執行緒處理待執行的動作
            while controller.running:
                try:
                    controller.process_pending_action()
                    # 檢查是否需要自動跳過換行符號
                    controller.check_and_skip_newline()
                except Exception as e:
                    logging.error(f"導航迴圈錯誤：{e}")
                    print(f"⚠️  操作發生錯誤：{e}")
                    controller.select_current_cell()
                time.sleep(0.05)  # 避免 CPU 佔用過高
        finally:
            if controller.listener:
                controller.listener.stop()

        # 【程式結束前】根據模式決定是否恢復人工標音文字顏色
        if not view_mode:
            print("\n正在恢復人工標音文字顏色...")
            restore_manual_annotation_style(wb)

        print("=" * 70)
        print("程式結束")
        print("=" * 70)
        return EXIT_CODE_SUCCESS

    except KeyError:
        print(f"錯誤：找不到工作表 '{SHEET_NAME}'")
        return EXIT_CODE_NO_FILE
    except Exception as e:
        logging.error(f"程式執行錯誤：{e}")
        # 發生錯誤時也要根據模式決定是否恢復樣式
        if not view_mode:
            try:
                restore_manual_annotation_style(wb)
            except Exception as e:
                logging.error(f"程式執行錯誤：{e}")
        return EXIT_CODE_ERROR


# =========================================================================
# 主要處理函數（使用輸入模式）
# =========================================================================
def read_han_ji_zu_im_sheet(wb, start_cell=None) -> int:
    """
    漢字注音工作表導讀主程式（輸入模式）

    Args:
        wb: Excel 工作簿物件
        start_cell: 校對起始儲存格（如 d133）；未指定時自 D5 開始

    Returns:
        退出代碼
    """
    try:
        # 取得工作表
        sheet = wb.sheets[SHEET_NAME]
        sheet.activate()

        # 取得總行數
        total_lines = get_total_lines(wb)

        try:
            current_row, current_col = resolve_start_cell(start_cell, total_lines)
        except ValueError as e:
            print(f"❌ {e}")
            return EXIT_CODE_INVALID_INPUT

        # 初始化：移動到指定起始儲存格（預設第一行行首 D5）
        sheet.range((current_row, current_col)).select()
        if start_cell:
            print(f"起始儲存格：{xw.utils.col_name(current_col)}{current_row}")

        print("=" * 70)
        print("漢字注音工作表導讀（輸入模式）")
        print("=" * 70)
        print("操作說明：")
        print("  ← (Left Arrow)  : 向左移動")
        print("  → (Right Arrow) : 向右移動")
        print("  pgup / pgdn     : 翻到上一頁／下一頁")
        print("  Ctrl+C          : 結束程式")
        print("=" * 70)
        print(f"總行數：{total_lines}")
        print(f"每行字數：{END_COL - START_COL + 1}")
        print("=" * 70)

        # 無限循環，等待使用者輸入
        while True:
            try:
                # 顯示當前位置
                line_no = get_line_number(current_row)
                col_letter = xw.utils.col_name(current_col)
                cell_value = sheet.range((current_row, current_col)).value or ""

                print(f"\n當前位置：第 {line_no} 行，儲存格 {col_letter}{current_row}【{cell_value}】")

                # 等待使用者輸入
                user_input = input("請按方向鍵（← / →）後按 Enter（Ctrl+C 結束）：").strip().lower()

                # 處理輸入
                if user_input in ["<-", "←", "left", "l"]:
                    # 向左移動
                    new_row, new_col = move_left(sheet, current_row, current_col)
                    if new_row != current_row or new_col != current_col:
                        current_row, current_col = new_row, new_col
                        sheet.range((current_row, current_col)).select()
                        print(f"→ 移動到：{xw.utils.col_name(current_col)}{current_row}")
                    else:
                        print("已在第一行行首，無法向左移動")

                elif user_input in ["->", "→", "right", "r"]:
                    # 向右移動
                    new_row, new_col = move_right(sheet, current_row, current_col, total_lines)
                    if new_row != current_row or new_col != current_col:
                        current_row, current_col = new_row, new_col
                        sheet.range((current_row, current_col)).select()
                        print(f"→ 移動到：{xw.utils.col_name(current_col)}{current_row}")
                    else:
                        print("已在最後一行行尾，無法向右移動")

                elif user_input in ["pgup", "pu"]:
                    # 翻到上一頁
                    lines_per_page = get_lines_per_page(wb)
                    line_no = get_line_number(current_row)
                    new_line = max(1, line_no - lines_per_page)
                    if new_line != line_no:
                        current_row = get_row_from_line(new_line)
                        sheet.range((current_row, current_col)).select()
                        print(f"→ 翻到上一頁：{xw.utils.col_name(current_col)}{current_row}")
                    else:
                        print("已在第一行，無法向上翻頁")

                elif user_input in ["pgdn", "pd"]:
                    # 翻到下一頁
                    lines_per_page = get_lines_per_page(wb)
                    line_no = get_line_number(current_row)
                    new_line = min(total_lines, line_no + lines_per_page)
                    if new_line != line_no:
                        current_row = get_row_from_line(new_line)
                        sheet.range((current_row, current_col)).select()
                        print(f"→ 翻到下一頁：{xw.utils.col_name(current_col)}{current_row}")
                    else:
                        print("已在最後一行，無法向下翻頁")

                elif user_input == "":
                    # 空白輸入，不移動
                    continue

                else:
                    print(f"無效的輸入：{user_input}")
                    print("請輸入：← (向左) 或 → (向右)")

            except KeyboardInterrupt:
                print("\n\n使用者中斷程式（Ctrl+C）")
                break
            except Exception as e:
                logging.error(f"處理錯誤：{e}")
                print(f"❌ 錯誤：{e}")
                continue

        print("=" * 70)
        print("程式結束")
        print("=" * 70)
        return EXIT_CODE_SUCCESS

    except Exception as e:
        logging.exception(f"程式執行失敗: {e}！")
        return EXIT_CODE_UNKNOWN_ERROR


def main(args) -> int:
    """主程式"""
    try:
        # 解析命令行參數
        view_mode = args.view
        start_cell = getattr(args, "start", None)

        # 在接觸 Excel 之前先記住啟動本程式的終端機（WezTerm 等）
        console_hwnd = capture_console_hwnd()

        # 取得 Excel 活頁簿
        wb = None
        # 若失敗，則取得作用中的活頁簿
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging.error(f"無法找到作用中的 Excel 工作簿: {e}")
            return EXIT_CODE_NO_FILE

        if not wb:
            logging.error("無法取得 Excel 活頁簿")
            return EXIT_CODE_NO_FILE

        # 根據是否安裝 pynput 決定使用哪種模式
        if HAS_PYNPUT:
            mode_text = "瀏覽模式" if view_mode else "校對模式"
            print(f"使用鍵盤監聽模式 - {mode_text}")
            return read_han_ji_with_keyboard(
                wb, view_mode=view_mode, console_hwnd=console_hwnd, start_cell=start_cell
            )
        else:
            print("使用輸入模式")
            return read_han_ji_zu_im_sheet(wb, start_cell=start_cell)

    except KeyboardInterrupt:
        print("\n\n使用者中斷程式（Ctrl+C）")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        logging.exception(f"程式執行失敗: {e}！")
        return EXIT_CODE_UNKNOWN_ERROR


# ==============================================================================
def test_01():
    """測試函數"""
    print("執行測試函數 test_01()")
    # 在這裡添加測試程式碼
    print("測試完成")


if __name__ == "__main__":
    import argparse
    import sys

    # 解析命令行參數
    parser = argparse.ArgumentParser(
        description="漢字注音工作表導讀程式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
python a300_漢字注音工作表校對.py                 # 校對模式，自 D5 開始
python a300_漢字注音工作表校對.py --view          # 瀏覽模式（隱藏人工標音）
python a300_漢字注音工作表校對.py --start d133    # 自儲存格 D133 續校
python a300_漢字注音工作表校對.py --start D133 --view
        """,
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="執行測試模式",
    )
    # parser.add_argument("--edit", action="store_true", help="啟用校稿模式（不隱藏人工標音文字顏色）")
    parser.add_argument(
        "--view",
        action="store_true",
        help="啟用瀏覽模式（隱藏人工標音文字顏色）",
    )
    parser.add_argument(
        "--start",
        metavar="CELL",
        default=None,
        type=argparse_start_cell,
        help="校對起始儲存格（例如 d133）。未指定時自 D5 開始。",
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
