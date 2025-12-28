# =========================================================================
# 程式功能摘要
# =========================================================================
# 用途：（1）使用者在瀏覽【漢字注音】工作表，逐字閱讀各【漢字儲存格】的【台語音標】
# 或【漢字標音】時，可用鍵盤【←】及【→】方向鍵，移動 Excel 的【作用儲存格】。
# 借助 Excel 的【作用儲存格】，以【行列十字交叉顯示】的特性，使用者能清楚聚焦
# 目光所在處。（2）當使用者遇到【台語音標】有問題之【漢字】，可立即按下【空白鍵】或【Q】鍵，
# 插斷正在執行的瀏覽動作，透過 a220_作用儲存格查找萌典漢字讀音.py 程式，進行漢字
# 讀音查詢工作；甚至決定將查詢的結果回寫【漢字注音】工作表，【作用儲存格】的
# 【人工標音儲存格】。
#
# 可利用【作用儲存格】十字交叉的聚焦游標，導引觀眾目光的移動，使逐字的漢字發音誦讀，
# 更顯有趣。另外，操作者無需借助滑鼠指標，僅需使用【←】或【→】按鍵，便能在上/下行
# 移動。譬如：在【第2行】的行尾（即儲存格：R9）時，按【→】鍵，游標會跳到【第3行】
# 的行首（即儲存格：D13）。
#
# 規格說明：
# （1）行號與【漢字儲存格】的對映關係如下：
#       -【第1行】的儲存格：D5, E5, F5,... ,R5；
#       -【第2行】的儲存格：D9, E9, F9,... ,R9；
#       -【第3行】的儲存格：D13, E13, F13,... ,R13；
# （2）程式開始執行時，【作用儲存格】落於【漢字注音】工作表的第1行行首儲存格（即：D5）。
# （3）操作鍵只提供【←】及【→】兩個方向鍵，分別用以控制游標向前或向後移動。
# （4）當【作用儲存格】游標位於某行的行首儲存格（如：D9、D13...）時，按【←】鍵，
#      游標會跳到前一行（如：D5、D9...）
# （5）當【作用儲存格】游標位於某行的行尾儲存格（如：R5、R9...）時，按【→】鍵，
#      游標會跳到下一行（如：D9、D13...）
# （6）若【作用儲存格】的下一個儲存格為【換行控制碼（\n）】，則【作用儲存格】的【聚焦游標】
#      須跳到下一行的行首儲存格。
# （7）若【作用儲存格】位於第一行的行首儲存格，按【←】鍵不動作。
# （8）若【作用儲存格】位於【行首儲存格】（如：D9, D13, D17...)，按【←】鍵，【聚焦游標】
#      需移到上一行；若上一行的【行尾】（如：R5, R9, R13...）有漢字或標點符號，在【作用儲存格】
#      落於行尾處；否則，需移到【換行】的前一個儲存格。
# （9）程式結束的按鍵為【ESC】或【Ctrl+C】鍵。
# （10）插斷瀏覽，進入漢字讀音查詢的按鍵為【空白鍵】或【Q】鍵。

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
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
    print("警告：未安裝 pywin32 套件，可能無法自動切換到 Excel 視窗")
    print("可執行：pip install pywin32")

# 載入自訂模組（用於字典查詢功能）
from mod_ChhoeTaigi import chhoe_taigi
from mod_excel_access import get_value_by_name
from mod_標音 import (
    PiauIm,
    convert_tl_with_tiau_hu_to_tlpa,
    format_han_ji_piau_im,
    split_tai_gi_im_piau,
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_NO_FILE = 1
EXIT_CODE_INVALID_INPUT = 2
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

        # 等待一下讓視窗切換完成
        time.sleep(0.3)

    except Exception as e:
        logging.error(f"無法激活 Excel 視窗：{e}")
        print(f"提示：請手動點擊 Excel 視窗以顯示十字游標")


def activate_console_window(console_hwnd):
    """
    激活終端機視窗，使其成為前景視窗

    Args:
        console_hwnd: 終端機視窗句柄
    """
    if not HAS_WIN32 or not console_hwnd:
        return

    try:
        # 檢查視窗是否存在
        if not win32gui.IsWindow(console_hwnd):
            print("無法找到終端機視窗")
            return

        # 將終端機視窗切換到前景
        win32gui.SetForegroundWindow(console_hwnd)

        # 等待一下讓視窗切換完成
        time.sleep(0.2)

    except Exception as e:
        logging.error(f"無法激活終端機視窗：{e}")


def query_han_ji_dictionary(sheet, row: int, col: int, piau_im: PiauIm, piau_im_huat: str):
    """
    查詢字典並讓使用者選擇讀音（獨立函數版本）

    Args:
        sheet: Excel 工作表物件
        row: 當前列號
        col: 當前欄號
        piau_im: 標音物件
        piau_im_huat: 標音方法
    """
    def _convert_piau_im(tai_lo_ping_im: str) -> tuple:
        """將台羅拼音轉換為音標"""
        # 將【台羅拼音】轉換成【台語音標】
        tlpa_im_piau = convert_tl_with_tiau_hu_to_tlpa(tai_lo_ping_im)
        # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
        siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tlpa_im_piau)

        # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
        tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])

        # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
        if (piau_im_huat == "十五音" or piau_im_huat == "雅俗通") and (siann_bu == "" or siann_bu == None):
            siann_bu = "ø"

        ok = False
        han_ji_piau_im = ""
        try:
            han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=piau_im_huat,
                siann_bu=siann_bu,
                un_bu=un_bu,
                tiau_ho=tiau_ho,
            )
            if han_ji_piau_im:
                ok = True
            else:
                logging.warning(f"【台語音標】：[{tai_gi_im_piau}]，轉換成【{piau_im_huat}漢字標音】拚音/注音系統失敗！")
        except Exception as e:
            logging.error(f"piau_im.han_ji_piau_im_tng_huan() 發生執行時期錯誤: 【台語音標】：{tai_gi_im_piau}, {e}")
            han_ji_piau_im = ""
            ok = False

        if not ok:
            return tai_gi_im_piau, ""
        else:
            return tai_gi_im_piau, format_han_ji_piau_im(han_ji_piau_im)

    try:
        # 取得當前儲存格的漢字
        cell = sheet.range((row, col))
        han_ji = cell.value

        if not han_ji or str(han_ji).strip() == "":
            print("當前儲存格沒有漢字")
            return

        han_ji = str(han_ji).strip()
        print(f"\n查詢漢字：【{han_ji}】")

        # 查詢萌典
        result = chhoe_taigi(han_ji=han_ji)

        # 查無此字
        if not result:
            print(f"【{han_ji}】查無此字！")
            return

        # 有多個讀音
        print(f"【{han_ji}】有 {len(result)} 個讀音：")

        # 顯示所有讀音選項
        piau_im_options = []
        for idx, tai_lo_ping_im in enumerate(result):
            # 轉換音標
            tai_gi_im_piau, han_ji_piau_im = _convert_piau_im(tai_lo_ping_im)
            piau_im_options.append((tai_gi_im_piau, han_ji_piau_im))
            msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
            print(f"{idx + 1}. {msg}")

        # 讓使用者選擇讀音
        user_input = input("\n請選擇讀音編號（直接按 Enter 略過，輸入編號後按 Enter 填入）：").strip()

        if user_input == "":
            # 只瀏覽，不填入
            print("略過填入，繼續導讀...")
            return

        try:
            choice = int(user_input)
            if 1 <= choice <= len(result):
                # 填入選擇的讀音
                tai_gi_im_piau, han_ji_piau_im = piau_im_options[choice - 1]
                cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
                cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
                cell.offset(1, 0).value = han_ji_piau_im    # 漢字標音
                print(f"✓ 已填入第 {choice} 個讀音：[{tai_gi_im_piau}] /【{han_ji_piau_im}】")
                print("繼續導讀...")
            else:
                print(f"無效的選擇：{choice}（超出範圍）")
        except ValueError:
            print(f"無效的輸入：{user_input}")

    except Exception as e:
        logging.error(f"字典查詢錯誤：{e}")
        print(f"查詢發生錯誤：{e}")


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

        # 記錄視窗句柄
        self.console_hwnd = None  # 終端機視窗句柄
        self.excel_hwnd = None    # Excel 視窗句柄
        if HAS_WIN32:
            try:
                # 記錄當前前景視窗（應該是終端機）
                self.console_hwnd = win32gui.GetForegroundWindow()
                # 記錄 Excel 視窗句柄
                self.excel_hwnd = wb.app.api.Hwnd
            except Exception as e:
                logging.error(f"無法記錄視窗句柄：{e}")

        # 初始化標音相關設定（用於字典查詢）
        han_ji_khoo_name = get_value_by_name(wb=wb, name='漢字庫')
        self.piau_im_huat = get_value_by_name(wb=wb, name='標音方法')
        self.piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)

    def move_to_cell(self, row, col):
        """移動到指定儲存格"""
        self.current_row = row
        self.current_col = col
        self.sheet.range((row, col)).select()

        # 顯示當前位置
        line_no = get_line_number(row)
        col_letter = xw.utils.col_name(col)
        cell_value = self.sheet.range((row, col)).value or ""
        print(f"→ 第 {line_no} 行，{col_letter}{row}【{cell_value}】")

    def on_key_press(self, key):
        """鍵盤按下事件處理 - 只設置動作標記"""
        try:
            if key == keyboard.Key.left:
                self.pending_action = 'left'
            elif key == keyboard.Key.right:
                self.pending_action = 'right'
            elif key == keyboard.Key.space:
                # 空白鍵：查詢字典
                self.pending_action = 'query'
            elif hasattr(key, 'char') and key.char and key.char.lower() == 'q':
                # Q 或 q 鍵：查詢字典
                self.pending_action = 'query'
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
                # 向左移動
                new_row, new_col = move_left(self.sheet, self.current_row, self.current_col)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == 'right':
                # 向右移動
                new_row, new_col = move_right(self.sheet, self.current_row, self.current_col, self.total_lines)
                if new_row != self.current_row or new_col != self.current_col:
                    self.move_to_cell(new_row, new_col)

            elif action == 'query':
                # 空白鍵或 Q 鍵：查詢字典
                print("\n進入字典查詢...")
                self.query_dictionary()

            elif action == 'esc':
                print("\n按下 ESC 鍵，程式結束")

        except Exception as e:
            logging.error(f"執行動作錯誤：{e}")

    def _convert_piau_im(self, tai_lo_ping_im: str) -> tuple:
        """
        將台羅拼音轉換為音標

        Args:
            tai_lo_ping_im: 台羅拼音

        Returns:
            (tai_gi_im_piau, han_ji_piau_im)
        """
        # 將【台羅拼音】轉換成【台語音標】
        tlpa_im_piau = convert_tl_with_tiau_hu_to_tlpa(tai_lo_ping_im)
        # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
        siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tlpa_im_piau)

        # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
        tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])

        # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
        if (self.piau_im_huat == "十五音" or self.piau_im_huat == "雅俗通") and (siann_bu == "" or siann_bu == None):
            siann_bu = "ø"

        ok = False
        han_ji_piau_im = ""
        try:
            han_ji_piau_im = self.piau_im.han_ji_piau_im_tng_huan(
                piau_im_huat=self.piau_im_huat,
                siann_bu=siann_bu,
                un_bu=un_bu,
                tiau_ho=tiau_ho,
            )
            if han_ji_piau_im:  # 傳回非空字串，表示【漢字標音】之轉換成功
                ok = True
            else:
                logging.warning(f"【台語音標】：[{tai_gi_im_piau}]，轉換成【{self.piau_im_huat}漢字標音】拚音/注音系統失敗！")
        except Exception as e:
            logging.error(f"piau_im.han_ji_piau_im_tng_huan() 發生執行時期錯誤: 【台語音標】：{tai_gi_im_piau}, {e}")
            han_ji_piau_im = ""
            ok = False

        # 若 ok 為 False，表示轉換失敗，則將【台語音標】直接傳回
        if not ok:
            return tai_gi_im_piau, ""
        else:
            return tai_gi_im_piau, format_han_ji_piau_im(han_ji_piau_im)

    def query_dictionary(self):
        """查詢字典並讓使用者選擇讀音"""
        # 暫停鍵盤監聽，讓 input() 能正常接收輸入
        if self.listener:
            print("\n[暫停鍵盤監聽，切換到輸入模式]")
            self.listener.stop()
            time.sleep(0.2)  # 等待監聽器完全停止

        # 切換到終端機視窗，讓輸入進入終端機而不是 Excel
        if self.console_hwnd:
            activate_console_window(self.console_hwnd)

        try:
            # 執行字典查詢
            query_han_ji_dictionary(
                sheet=self.sheet,
                row=self.current_row,
                col=self.current_col,
                piau_im=self.piau_im,
                piau_im_huat=self.piau_im_huat
            )
        finally:
            # 切換回 Excel 視窗
            if self.excel_hwnd:
                activate_excel_window(self.wb)

            # 重新啟動鍵盤監聽
            if self.listener:
                print("\n[重新啟動鍵盤監聽]")
                self.listener = keyboard.Listener(
                    on_press=self.on_key_press,
                    suppress=True
                )
                self.listener.start()
                time.sleep(0.2)  # 等待監聽器啟動
                print("✓ 繼續使用方向鍵導航...\n")


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

        print("="  * 70)
        print("漢字注音工作表導讀（鍵盤監聽模式）")
        print("=" * 70)
        print("操作說明：")
        print("  ← (Left Arrow)  : 向左移動")
        print("  → (Right Arrow) : 向右移動")
        print("  Space / Q       : 查詢字典（查詢當前漢字讀音）")
        print("  ESC             : 結束程式")
        print("=" * 70)
        print(f"總行數：{controller.total_lines}")
        print(f"每行字數：{END_COL - START_COL + 1}")
        print("=" * 70)

        # 切換到 Excel 視窗，讓十字游標顯示
        print("\n正在切換到 Excel 視窗...")
        activate_excel_window(wb)

        print("\n請使用方向鍵導航...")
        print("提示：程式會攔截按鍵，不會影響 Excel 儲存格內容")

        # 啟動鍵盤監聽（在背景執行緒）
        # suppress=True 會阻止按鍵事件傳遞到其他應用程序（如 Excel）
        controller.listener = keyboard.Listener(
            on_press=controller.on_key_press,
            suppress=True  # 關鍵：阻止按鍵傳遞到 Excel
        )
        controller.listener.start()

        try:
            # 主迴圈：在主執行緒處理待執行的動作
            while controller.running:
                controller.process_pending_action()
                time.sleep(0.05)  # 避免 CPU 佔用過高
        finally:
            if controller.listener:
                controller.listener.stop()

        print("=" * 70)
        print("程式結束")
        print("=" * 70)
        return EXIT_CODE_SUCCESS

    except KeyError:
        print(f"錯誤：找不到工作表 '{SHEET_NAME}'")
        return EXIT_CODE_NO_FILE
    except Exception as e:
        logging.error(f"程式執行錯誤：{e}")
        return EXIT_CODE_UNKNOWN_ERROR


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

        # 初始化標音相關設定（用於字典查詢）
        han_ji_khoo_name = get_value_by_name(wb=wb, name='漢字庫')
        piau_im_huat = get_value_by_name(wb=wb, name='標音方法')
        piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)

        # 初始化：移動到第一行行首（D5）
        current_row = START_ROW
        current_col = START_COL
        sheet.range((current_row, current_col)).select()

        print("="  * 70)
        print("漢字注音工作表導讀（輸入模式）")
        print("=" * 70)
        print("操作說明：")
        print("  ← (Left Arrow)  : 向左移動")
        print("  → (Right Arrow) : 向右移動")
        print("  Space / Q       : 查詢字典（查詢當前漢字讀音）")
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
                user_input = input("請輸入：← / → (移動)、Space/Q (查字典)、Enter (Ctrl+C 結束)：").strip().lower()

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

                elif user_input in ['q', ' ', 'space']:
                    # 空白鍵或 Q 鍵：查詢字典
                    print("\n進入字典查詢...")
                    query_han_ji_dictionary(
                        sheet=sheet,
                        row=current_row,
                        col=current_col,
                        piau_im=piau_im,
                        piau_im_huat=piau_im_huat
                    )

                elif user_input == '':
                    # 空白輸入，不移動
                    continue

                else:
                    print(f"無效的輸入：{user_input}")
                    print("請輸入：← (向左)、→ (向右)、Space/Q (查字典)")

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
