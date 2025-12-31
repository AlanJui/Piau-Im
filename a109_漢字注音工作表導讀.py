# =========================================================================
# 程式功能摘要
# =========================================================================
# 用途：提供 <-- 及 --> （向前/向後）按鍵，以利操作者在誦讀【漢字注音】工作表時，
# 可利用【作用儲存格】十字交叉的聚焦游標，導引觀眾目光的移動，使逐字的漢字發音誦讀，
# 更顯有趣。另外，操作者無需借助滑鼠指標，僅需使用【←】或【→】按鍵，便能在上/下行
# 移動。譬如：在【第2行】的行尾（即儲存格：R9）時，按【→】鍵，游標會跳到【第3行】
# 的行首（即儲存格：D13）。

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
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
    import pythoncom
    HAS_PYTHONCOM = True
except ImportError:
    HAS_PYTHONCOM = False

# 載入 a222 的核心查詢功能（個人字典）
try:
    from a222_依作用儲存格在個人字典查找漢字讀音 import ca_han_ji_thak_im as ca_han_ji_thak_im_a222
    HAS_A222 = True
except ImportError as e:
    HAS_A222 = False
    print(f"警告：無法載入 a222 模組：{e}")

# 載入 a220 的核心查詢功能（萌典）
try:
    from a220_作用儲存格查找萌典漢字讀音 import ca_han_ji_thak_im as ca_han_ji_thak_im_a220
    HAS_A220 = True
except ImportError as e:
    HAS_A220 = False
    print(f"警告：無法載入 a220 模組：{e}")

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_ERROR = 10
EXIT_CODE_UNKNOWN_ERROR = 99

# 工作表設定
SHEET_NAME = '漢字注音'
START_ROW = 5       # 第一行的起始列號
START_COL = 4       # D 欄（第 4 欄）
END_COL = 18        # R 欄（第 18 欄）
ROWS_PER_LINE = 4   # 每行佔用 4 列

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


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
                if cell_formula and '=CHAR(10)' in cell_formula.upper():
                    continue
                if cell_value == '\n':
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
    if next_cell_formula and '=CHAR(10)' in next_cell_formula.upper():
        is_newline = True
        # print(f"    → 偵測到 CHAR(10) 公式")

    # 方法2: 檢查值是否為換行符
    elif next_cell_value is not None:
        if next_cell_value == '\n' or next_cell_value == chr(10):
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
        total_lines = int(wb.names['每頁總列數'].refers_to_range.value)
        return total_lines
    except:
        # 預設值
        return 10


def hide_manual_annotation_style(wb):
    """
    隱藏【人工標音儲存格】樣式的文字
    將字型顏色改為與填滿顏色相同（象牙白）

    Args:
        wb: Excel 工作簿物件
    """
    try:
        # 取得 Excel API 物件
        excel_app = wb.app.api
        workbook = wb.api

        # 查找【人工標音儲存格】樣式
        style_name = "人工標音儲存格"
        try:
            style = workbook.Styles(style_name)
            # 將字型顏色改為象牙白（RGB: 255, 255, 240）
            # Excel 使用 BGR 格式，所以順序相反
            style.Font.Color = 0xF0FFFF  # BGR: 240, 255, 255 (象牙白)
            print(f"✓ 已隱藏【{style_name}】樣式的文字（字型顏色改為象牙白）")
        except:
            print(f"⚠️  找不到【{style_name}】樣式，跳過隱藏操作")

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
        excel_app = wb.app.api
        workbook = wb.api

        # 查找【人工標音儲存格】樣式
        style_name = "人工標音儲存格"
        try:
            style = workbook.Styles(style_name)
            # 將字型顏色改回紅色（RGB: 255, 0, 0）
            # Excel 使用 BGR 格式，所以順序相反
            style.Font.Color = 0x0000FF  # BGR: 0, 0, 255 (紅色)
            print(f"✓ 已恢復【{style_name}】樣式的文字（字型顏色改回紅色）")
        except:
            print(f"⚠️  找不到【{style_name}】樣式，跳過恢復操作")

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


def activate_console_window(console_hwnd):
    """
    激活終端機視窗，使其成為前景視窗

    Args:
        console_hwnd: 終端機視窗句柄
    """
    if not HAS_WIN32:
        print("提示：無法自動切換到終端機視窗（需要 pywin32 套件）")
        return

    try:
        import win32api
        import win32process

        # 嘗試找到正確的 Console 視窗
        current_hwnd = console_hwnd

        # 如果提供的句柄無效，嘗試找到 Python 控制台或 PowerShell 視窗
        if not current_hwnd or not win32gui.IsWindow(current_hwnd):
            def enum_handler(hwnd, result_list):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if any(keyword in title.lower() for keyword in ['python', 'powershell', 'cmd', 'terminal', 'piau-im', 'vscode']):
                        result_list.append(hwnd)

            windows = []
            win32gui.EnumWindows(enum_handler, windows)
            if windows:
                current_hwnd = windows[0]

        if current_hwnd and win32gui.IsWindow(current_hwnd):
            # 如果視窗最小化，先還原
            if win32gui.IsIconic(current_hwnd):
                win32gui.ShowWindow(current_hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)

            # 【強化版】使用 AttachThreadInput 解決 Windows 前景視窗限制
            try:
                # 獲取當前前景視窗的線程ID
                current_thread_id = win32api.GetCurrentThreadId()
                # 獲取目標視窗的線程ID
                target_thread_id, _ = win32process.GetWindowThreadProcessId(current_hwnd)

                # 如果線程不同，嘗試附加線程輸入
                if current_thread_id != target_thread_id:
                    try:
                        win32process.AttachThreadInput(current_thread_id, target_thread_id, True)
                        logging.debug(f"成功附加線程輸入: {current_thread_id} -> {target_thread_id}")
                    except Exception as e:
                        logging.debug(f"AttachThreadInput 失敗: {e}")

                # 方法 1: 使用 BringWindowToTop
                win32gui.BringWindowToTop(current_hwnd)
                time.sleep(0.1)

                # 方法 2: 使用 ShowWindow 激活
                win32gui.ShowWindow(current_hwnd, win32con.SW_SHOW)
                time.sleep(0.1)

                # 方法 3: 設為前景視窗
                win32gui.SetForegroundWindow(current_hwnd)
                time.sleep(0.3)

                # 方法 4: 再次嘗試激活
                win32gui.SetActiveWindow(current_hwnd)

                # 分離線程輸入
                if current_thread_id != target_thread_id:
                    try:
                        win32process.AttachThreadInput(current_thread_id, target_thread_id, False)
                    except Exception as e:
                        logging.debug(f"DetachThreadInput 失敗: {e}")

            except Exception as e:
                # SetActiveWindow 可能失敗，這是正常的
                logging.debug(f"視窗激活過程出現錯誤（可預期）：{e}")

            print("✓ 已切換到終端機視窗")

            # 等待更長時間確保視窗完全激活並準備接收輸入
            time.sleep(1.0)

            # 驗證視窗是否成為前景視窗
            foreground = win32gui.GetForegroundWindow()
            if foreground != current_hwnd:
                print(f"⚠️  視窗切換可能未完成")
                print(f"提示：請用滑鼠點擊一次終端機視窗以確保輸入焦點正確")
            else:
                # 即使前景視窗正確，仍然建議用戶確認
                print(f"提示：如果無法輸入，請用滑鼠點擊一次終端機視窗")
        else:
            print("提示：無法找到終端機視窗，請手動點擊終端機視窗")
    except Exception as e:
        # Windows 對 SetForegroundWindow 有限制，可能會失敗
        # 這不是致命錯誤，只需提示用戶手動點擊
        print(f"提示：無法自動切換視窗，請手動點擊終端機視窗")
        logging.debug(f"SetForegroundWindow 失敗：{e}")


# =========================================================================
# 主要處理函數（使用鍵盤監聽）
# =========================================================================
class NavigationController:
    """導航控制器 - 使用鍵盤監聽"""

    def __init__(self, wb, sheet):
        self.wb = wb
        self.sheet = sheet
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
        self.console_hwnd = None
        self.excel_hwnd = None
        if HAS_WIN32:
            try:
                # 取得 Excel 視窗句柄（使用 xlwings API）
                self.excel_hwnd = wb.app.api.Hwnd

                # 取得當前前景視窗（應該是 Console）
                current_foreground = win32gui.GetForegroundWindow()

                # 驗證這是否是 Console 視窗
                if current_foreground:
                    title = win32gui.GetWindowText(current_foreground)
                    # 如果標題包含 Python, PowerShell, CMD 等，這就是 Console
                    if any(keyword in title.lower() for keyword in ['python', 'powershell', 'cmd', 'terminal', 'piau-im', 'vscode']):
                        self.console_hwnd = current_foreground
                    else:
                        # 否則嘗試搜尋 Console 視窗
                        self.console_hwnd = self._find_console_window()

                logging.info(f"Console 視窗句柄：{self.console_hwnd}")
                logging.info(f"Excel 視窗句柄：{self.excel_hwnd}")
            except Exception as e:
                logging.warning(f"無法取得視窗句柄：{e}")

    def _find_console_window(self):
        """搜尋 Console 視窗"""
        try:
            windows = []
            def enum_handler(hwnd, result_list):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if any(keyword in title.lower() for keyword in ['python', 'powershell', 'cmd', 'terminal', 'piau-im', 'vscode']):
                        result_list.append(hwnd)

            win32gui.EnumWindows(enum_handler, windows)
            return windows[0] if windows else None
        except Exception as e:
            logging.warning(f"搜尋 Console 視窗失敗：{e}")
            return None

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
        current_cell = self.sheet.range((self.current_row, self.current_col))
        cell_value = current_cell.value
        cell_formula = current_cell.formula

        is_newline = False
        # 檢查公式是否為 =CHAR(10)
        if cell_formula and '=CHAR(10)' in str(cell_formula).upper():
            is_newline = True
        # 檢查值是否為換行符
        elif cell_value is not None:
            if cell_value == '\n' or cell_value == chr(10):
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
                self.pending_action = 'left'
            elif key == keyboard.Key.right:
                self.pending_action = 'right'
            elif key == keyboard.Key.up:
                self.pending_action = 'up'
            elif key == keyboard.Key.down:
                self.pending_action = 'down'
            elif key == keyboard.Key.space:
                # 空白鍵：查詢萌典
                self.pending_action = 'query_moedict'
            elif hasattr(key, 'char') and key.char:
                # 處理字元鍵
                if key.char.lower() == 'q':
                    # Q 鍵：查詢萌典
                    self.pending_action = 'query_moedict'
                elif key.char.lower() == 's':
                    # S 鍵：查詢個人字典
                    self.pending_action = 'query_personal'
            elif key == keyboard.Key.esc:
                self.pending_action = 'esc'
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
            if action == 'left':
                # 向左移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_left(self.sheet, self.current_row, self.current_col)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == 'right':
                # 向右移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_right(self.sheet, self.current_row, self.current_col, self.total_lines)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == 'up':
                # 向上移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_up(self.sheet, self.current_row, self.current_col)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == 'down':
                # 向下移動（重置延遲計時器）
                self.last_move_time = None
                new_row, new_col = move_down(self.sheet, self.current_row, self.current_col, self.total_lines)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == 'query_moedict':
                # 查詢萌典
                self.query_moedict_dictionary()

            elif action == 'query_personal':
                # 查詢個人字典
                self.query_personal_dictionary()

            elif action == 'esc':
                print("\n按下 ESC 鍵，程式結束")

        except Exception as e:
            logging.error(f"執行動作錯誤：{e}")

    def query_moedict_dictionary(self):
        """查詢萌典字典"""
        print("\n" + "=" * 70)
        print("進入萌典字典查詢模式")
        print("=" * 70)

        # 暫停鍵盤監聽
        if self.listener:
            self.listener.stop()
            time.sleep(0.3)

        try:
            if HAS_A220:
                # 直接調用 a220 的核心函數，不進入無限循環
                print("\n查詢萌典字典中...")

                # 切換到終端機視窗（確保用戶可以輸入）
                activate_console_window(self.console_hwnd)

                # 取得設定值
                try:
                    from mod_excel_access import get_value_by_name
                    ue_im_lui_piat = get_value_by_name(wb=self.wb, name='語音類型')
                    han_ji_khoo = get_value_by_name(wb=self.wb, name='漢字庫')
                except:
                    ue_im_lui_piat = "白話音"
                    han_ji_khoo = "河洛話"

                # 取得當前作用儲存格位置
                current_cell = f"{xw.utils.col_name(self.current_col)}{self.current_row}"
                print(f"當前儲存格：{current_cell}")

                # 調用查詢函數
                exit_code = ca_han_ji_thak_im_a220(
                    wb=self.wb,
                    sheet_name='漢字注音',
                    cell=current_cell,
                    ue_im_lui_piat=ue_im_lui_piat,
                    han_ji_khoo=han_ji_khoo,
                    new_khuat_ji_piau_sheet=False,
                    new_piau_im_ji_khoo_sheet=False,
                )

                if exit_code == 0:
                    print("\n✓ 查詢完成")
                else:
                    print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
            else:
                # 回退到 subprocess 方式
                print("\n執行 a220_作用儲存格查找萌典漢字讀音.py...")
                result = subprocess.run(
                    [sys.executable, "a220_作用儲存格查找萌典漢字讀音.py"],
                    cwd=os.path.dirname(os.path.abspath(__file__)),
                    capture_output=False,
                    text=True
                )
                if result.returncode != 0:
                    print(f"⚠️  a220 程式執行失敗，返回碼：{result.returncode}")
        except KeyboardInterrupt:
            print("\n\n使用者中斷查詢")
        except Exception as e:
            logging.error(f"執行萌典查詢失敗：{e}")
            print(f"❌ 執行萌典查詢失敗：{e}")
        finally:
            print("\n" + "=" * 70)
            print("返回導航模式")
            print("=" * 70)

            # 切換回 Excel 視窗
            activate_excel_window(self.wb)

            # 重新啟動鍵盤監聽
            if self.listener:
                self.listener = keyboard.Listener(
                    on_press=self.on_key_press,
                    suppress=True
                )
                self.listener.start()
                time.sleep(0.3)
            print("✓ 已恢復導航模式\n")

    def query_personal_dictionary(self):
        """查詢個人字典"""
        print("\n" + "=" * 70)
        print("進入個人字典查詢模式")
        print("=" * 70)

        # 暫停鍵盤監聽
        if self.listener:
            self.listener.stop()
            time.sleep(0.3)

        try:
            if HAS_A222:
                # 直接調用 a222 的核心函數，不進入無限循環
                print("\n查詢個人字典中...")

                # 切換到終端機視窗（確保用戶可以輸入）
                activate_console_window(self.console_hwnd)

                # 取得設定值
                try:
                    from mod_excel_access import get_value_by_name
                    ue_im_lui_piat = get_value_by_name(wb=self.wb, name='語音類型')
                    han_ji_khoo = get_value_by_name(wb=self.wb, name='漢字庫')
                except:
                    ue_im_lui_piat = "白話音"
                    han_ji_khoo = "河洛話"

                # 取得當前作用儲存格位置
                current_cell = f"{xw.utils.col_name(self.current_col)}{self.current_row}"
                print(f"當前儲存格：{current_cell}")

                # 調用查詢函數
                exit_code = ca_han_ji_thak_im_a222(
                    wb=self.wb,
                    sheet_name='漢字注音',
                    cell=current_cell,
                    ue_im_lui_piat=ue_im_lui_piat,
                    han_ji_khoo=han_ji_khoo,
                    new_khuat_ji_piau_sheet=False,
                    new_piau_im_ji_khoo_sheet=False,
                )

                if exit_code == 0:
                    print("\n✓ 查詢完成")
                else:
                    print(f"\n⚠️  查詢結果：exit_code = {exit_code}")
            else:
                # 回退到 subprocess 方式
                print("\n執行 a222_依作用儲存格在個人字典查找漢字讀音.py...")
                result = subprocess.run(
                    [sys.executable, "a222_依作用儲存格在個人字典查找漢字讀音.py"],
                    cwd=os.path.dirname(os.path.abspath(__file__)),
                    capture_output=False,
                    text=True
                )
                if result.returncode != 0:
                    print(f"⚠️  a222 程式執行失敗，返回碼：{result.returncode}")
        except KeyboardInterrupt:
            print("\n\n使用者中斷查詢")
        except Exception as e:
            logging.error(f"執行個人字典查詢失敗：{e}")
            print(f"❌ 執行個人字典查詢失敗：{e}")
        finally:
            print("\n" + "=" * 70)
            print("返回導航模式")
            print("=" * 70)

            # 切換回 Excel 視窗
            activate_excel_window(self.wb)

            # 重新啟動鍵盤監聽
            if self.listener:
                self.listener = keyboard.Listener(
                    on_press=self.on_key_press,
                    suppress=True
                )
                self.listener.start()
                time.sleep(0.3)
            print("✓ 已恢復導航模式\n")


def read_han_ji_with_keyboard(wb) -> int:
    """
    漢字注音工作表導讀主程式（使用鍵盤監聽）

    Args:
        wb: Excel 工作簿物件

    Returns:
        退出代碼
    """
    try:
        # 取得工作表
        sheet = wb.sheets[SHEET_NAME]
        sheet.activate()

        # 初始化控制器
        controller = NavigationController(wb, sheet)

        # 移動到第一行行首（D5）
        controller.move_to_cell(START_ROW, START_COL)

        print("=" * 70)
        print("漢字注音工作表導讀（鍵盤監聽模式）")
        print("=" * 70)
        print("操作說明：")
        print("  ← (Left Arrow)  : 向左移動")
        print("  → (Right Arrow) : 向右移動")
        print("  ↑ (Up Arrow)    : 向上移動到上一行")
        print("  ↓ (Down Arrow)  : 向下移動到下一行")
        print("  空白 / Q 鍵     : 查詢萌典字典")
        print("  S 鍵            : 查詢個人字典")
        print("  ESC             : 結束程式")
        print("=" * 70)
        print(f"總行數：{controller.total_lines}")
        print(f"每行字數：{END_COL - START_COL + 1}")
        print("=" * 70)

        # 【進入導讀模式前】隱藏人工標音文字
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
            suppress=True  # 攔截按鍵，不讓 Excel 接收
        )
        controller.listener.start()

        try:
            # 主迴圈：在主執行緒處理待執行的動作
            while controller.running:
                controller.process_pending_action()
                # 檢查是否需要自動跳過換行符號
                controller.check_and_skip_newline()
                time.sleep(0.05)  # 避免 CPU 佔用過高
        finally:
            if controller.listener:
                controller.listener.stop()

        # 【程式結束前】恢復人工標音文字顏色
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
        # 發生錯誤時也要恢復樣式
        try:
            restore_manual_annotation_style(wb)
        except:
            pass
        return EXIT_CODE_ERROR


# =========================================================================
# 主要處理函數（使用輸入模式）
# =========================================================================
def read_han_ji_zu_im_sheet(wb) -> int:
    """
    漢字注音工作表導讀主程式（輸入模式）

    Args:
        wb: Excel 工作簿物件

    Returns:
        退出代碼
    """
    try:
        # 取得工作表
        sheet = wb.sheets[SHEET_NAME]
        sheet.activate()

        # 取得總行數
        total_lines = get_total_lines(wb)

        # 初始化：移動到第一行行首（D5）
        current_row = START_ROW
        current_col = START_COL
        sheet.range((current_row, current_col)).select()

        print("=" * 70)
        print("漢字注音工作表導讀（輸入模式）")
        print("=" * 70)
        print("操作說明：")
        print("  ← (Left Arrow)  : 向左移動")
        print("  → (Right Arrow) : 向右移動")
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
                if user_input in ['<-', '←', 'left', 'l']:
                    # 向左移動
                    new_row, new_col = move_left(sheet, current_row, current_col)
                    if new_row != current_row or new_col != current_col:
                        current_row, current_col = new_row, new_col
                        sheet.range((current_row, current_col)).select()
                        print(f"→ 移動到：{xw.utils.col_name(current_col)}{current_row}")
                    else:
                        print("已在第一行行首，無法向左移動")

                elif user_input in ['->', '→', 'right', 'r']:
                    # 向右移動
                    new_row, new_col = move_right(sheet, current_row, current_col, total_lines)
                    if new_row != current_row or new_col != current_col:
                        current_row, current_col = new_row, new_col
                        sheet.range((current_row, current_col)).select()
                        print(f"→ 移動到：{xw.utils.col_name(current_col)}{current_row}")
                    else:
                        print("已在最後一行行尾，無法向右移動")

                elif user_input == '':
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
        logging.exception("程式執行失敗")
        return EXIT_CODE_UNKNOWN_ERROR


def main():
    """主程式"""
    try:
        # 取得 Excel 活頁簿
        wb = None
        try:
            # 嘗試從 Excel 呼叫取得（RunPython）
            wb = xw.Book.caller()
        except:
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
            print("使用鍵盤監聽模式")
            return read_han_ji_with_keyboard(wb)
        else:
            print("使用輸入模式")
            return read_han_ji_zu_im_sheet(wb)

    except KeyboardInterrupt:
        print("\n\n使用者中斷程式（Ctrl+C）")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        logging.exception("程式執行失敗")
        return EXIT_CODE_UNKNOWN_ERROR


if __name__ == "__main__":
    import sys
    sys.exit(main())
