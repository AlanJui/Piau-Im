# =========================================================================
# 程式功能摘要
# =========================================================================
# 用途：提供 <-- 及 --> （向前/向後）按鍵，以利操作者在誦讀【漢字注音】工作表時，
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
            # 跳到前一行的行尾
            new_line = line_no - 1
            new_row = get_row_from_line(new_line)
            new_col = END_COL
            return new_row, new_col
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

            elif action == 'esc':
                print("\n按下 ESC 鍵，程式結束")

        except Exception as e:
            logging.error(f"執行動作錯誤：{e}")


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
        print("  ESC             : 結束程式")
        print("=" * 70)
        print(f"總行數：{controller.total_lines}")
        print(f"每行字數：{END_COL - START_COL + 1}")
        print("=" * 70)
        print("\n請使用方向鍵導航...")

        # 啟動鍵盤監聽（在背景執行緒）
        listener = keyboard.Listener(on_press=controller.on_key_press)
        listener.start()

        try:
            # 主迴圈：在主執行緒處理待執行的動作
            while controller.running:
                controller.process_pending_action()
                time.sleep(0.05)  # 避免 CPU 佔用過高
        finally:
            listener.stop()

        print("=" * 70)
        print("程式結束")
        print("=" * 70)
        return EXIT_CODE_SUCCESS

    except KeyError:
        print(f"錯誤：找不到工作表 '{SHEET_NAME}'")
        return EXIT_CODE_NO_FILE
    except Exception as e:
        logging.error(f"程式執行錯誤：{e}")
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
