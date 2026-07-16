"""
a330_在字庫工作表查找漢字讀音.py v0.0.6
功能說明：
    在【缺字表/標音字庫/人工標音字庫】工作表，選定某【資料紀錄】，使期成為【作用儲存格】。
    本函式便會自【A欄】取得【漢字】，在【B欄】取得【台語音標】，然後在【漢字庫】字典查找
    漢字讀音，或由使用者輸入【台語音標】。最後將【台語音標】回填【C欄】【校正音標】。
操作說明（鍵盤監聽模式，參考 a109 作法）：
    1. 自 Terminal 啟動程式（如：py a340.py），操作提示顯示後，視窗聚焦自動切回 Excel；
    2. 以滑鼠點選字庫工作表之儲存格，指定【作用儲存格】；
    3. 按 <Enter> 或 <Space> 鍵：執行【漢字查找讀音】作業，將查得/輸入之讀音填入
       C 欄【校正音標】；作業結束後，視窗聚焦自動切回 Excel；
    4. 按 <Esc> 鍵：終止程式（並自動存檔）。
    （若未安裝 pynput 套件，則退回【輸入模式】：於 Terminal 按 Enter 查詢、Ctrl+C 終止。）
更新紀錄：
v0.0.6: 新增 6 個導航快捷鍵（因監聽器以 suppress=True 攔截按鍵，Excel 無法直接
        接收，故由程式代收後以 xlwings 執行等效操作）：
        - ↑/↓：作用儲存格上／下移一列；
        - PgUp/PgDn：作用儲存格上／下移一個捲頁；
        - Ctrl+PgUp/PgDn：切換至前／後一個工作表。
v0.0.5: （1）程式啟動、操作提示顯示完畢後，視窗聚焦自動切回 Excel；（2）<Space> 鍵
        等同 <Enter> 鍵，均可觸發查音作業；（3）查音作業結束後，視窗聚焦自動切回
        Excel（強化：Excel 視窗切換改用 AttachThreadInput，解決 Windows 前景視窗
        限制，使切換更可靠）。
v0.0.4: 參考 a109 作法，改用 pynput 鍵盤監聽模式：按 <Enter> 執行查音、按 <Esc> 終止
        程式。並修正原【輸入模式】按 Ctrl+C 無法終止程式之問題（原程式於 exit_code
        非 0 時，捕捉 KeyboardInterrupt 後未返回，致迴圈繼續執行）。
v0.0.3: 自 a221 變更成 a340 。
v0.0.2: 不再僅限於單一工作表使用；依【作用儲存格】所在之工作表（缺字表/標音字庫/
        人工標音字庫）自動判別作業模式：
        - 缺字表：查得【台語音標】後回填 B 欄，並將紀錄移轉至【標音字庫】；
        - 標音字庫/人工標音字庫：查得【台語音標】後回填 C 欄【校正音標】。
        兩種模式均依 D 欄【座標】，回填【漢字注音】工作表之【台語音標】與【漢字標音】。
v0.0.1: 自 a222.py 改寫。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import time
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# 鍵盤監聽套件（參考 a109 作法）
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

# 載入自訂模組
from mod_excel_access import (
    convert_coord_list_str_to_tuples,
    excel_address_to_row_col,
    get_active_cell_address,
)
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
    logging_process_step,  # noqa: F401
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

# 本程式可作用之【字庫】工作表名稱
KHUAT_JI_PIAU_SHEET = "缺字表"
PIAU_IM_JI_KHOO_SHEET = "標音字庫"
JIN_KANG_PIAU_IM_JI_KHOO_SHEET = "人工標音字庫"
VALID_JI_KHOO_SHEETS = (
    KHUAT_JI_PIAU_SHEET,
    PIAU_IM_JI_KHOO_SHEET,
    JIN_KANG_PIAU_IM_JI_KHOO_SHEET,
)

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
    def _update_piau_im_ji_khoo_from_khuat_ji_piau(
        self,
        han_ji: str,
        tai_gi_im_piau: str,
        coordinates: list[tuple],
    ) -> None:
        """
        將原本在【缺字表】工作表之【資料紀錄】，寫入【標音字庫】工作表，以示該漢字已查得
        【台語音標】，並已在【漢字注音】工作表完成【台語音標】及【漢字標音】之標注工作。

        處理步驟：
        （1）依【座標清單】，將【漢字＋台語音標＋座標】逐筆登錄至【標音字庫】字典；
        （2）自【缺字表】字典，逐一移除該漢字之【座標】（座標清空時，整筆紀錄一併移除）；
        （3）將上述兩字典之內容，回寫至【標音字庫】與【缺字表】工作表。

        Args:
            han_ji: 已查得讀音之漢字
            tai_gi_im_piau: 自個人字典查得之台語音標
            coordinates: 指向【漢字注音】工作表【漢字】儲存格之座標清單
        """
        wb = self.program.wb

        for coord in coordinates:
            # 座標若為字串（re.findall 解析結果），先轉換成整數座標
            coordinate = (int(coord[0]), int(coord[1]))

            # （1）在【標音字庫】工作表（字庫物件），新增或更新一筆【資料紀錄】
            self.piau_im_ji_khoo_dict.add_or_update_entry(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                hau_ziann_im_piau="N/A",
                coordinates=coordinate,
            )

            # （2）自【缺字表】字典物件，移除該漢字之此一【座標】
            self.khuat_ji_piau_ji_khoo_dict.remove_coordinate_by_han_ji_and_coordinate(
                han_ji=han_ji,
                coordinate=coordinate,
            )

        # （3）將兩【字典物件】之內容，回寫至各自之工作表
        self.piau_im_ji_khoo_dict.save_to_sheet(
            wb=wb,
            sheet_name=self.piau_im_ji_khoo_dict.name,
        )
        self.khuat_ji_piau_ji_khoo_dict.save_to_sheet(
            wb=wb,
            sheet_name=self.khuat_ji_piau_ji_khoo_dict.name,
        )


# =========================================================================
# 視窗切換函數（參考 a109 作法）
# =========================================================================
def _find_console_window():
    """搜尋 Console（Terminal）視窗"""
    try:
        windows = []

        def enum_handler(hwnd, result_list):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if any(
                    keyword in title.lower()
                    for keyword in [
                        "python",
                        "powershell",
                        "cmd",
                        "terminal",
                        "piau-im",
                        "vscode",
                    ]
                ):
                    result_list.append(hwnd)

        win32gui.EnumWindows(enum_handler, windows)
        return windows[0] if windows else None
    except Exception as e:
        logging.warning(f"搜尋 Console 視窗失敗：{e}")
        return None


def _force_foreground_window(hwnd) -> bool:
    """
    強制將指定視窗設為前景視窗。

    Windows 對 SetForegroundWindow 有限制（背景程式不得任意搶奪前景），
    故採 AttachThreadInput 附加線程輸入後，多次嘗試激活視窗。

    Returns:
        True: 已成功成為前景視窗；False: 切換可能未完成
    """
    try:
        import win32process

        # 如果視窗最小化，先還原
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)

        # 使用 AttachThreadInput 解決 Windows 前景視窗限制
        foreground_hwnd = win32gui.GetForegroundWindow()
        foreground_thread_id, _ = win32process.GetWindowThreadProcessId(foreground_hwnd)
        target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)

        attached = False
        if foreground_thread_id != target_thread_id:
            try:
                win32process.AttachThreadInput(foreground_thread_id, target_thread_id, True)
                attached = True
            except Exception as e:
                logging.debug(f"AttachThreadInput 失敗: {e}")

        # 多次嘗試激活視窗
        for _ in range(3):
            win32gui.BringWindowToTop(hwnd)
            time.sleep(0.1)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            time.sleep(0.1)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)

            if win32gui.GetForegroundWindow() == hwnd:
                break
            time.sleep(0.2)

        try:
            win32gui.SetActiveWindow(hwnd)
        except Exception as e:
            logging.debug(f"SetActiveWindow 失敗: {e}！")

        # 分離線程輸入
        if attached:
            try:
                win32process.AttachThreadInput(foreground_thread_id, target_thread_id, False)
            except Exception as e:
                logging.debug(f"DetachThreadInput 失敗: {e}")

        return win32gui.GetForegroundWindow() == hwnd

    except Exception as e:
        logging.debug(f"視窗激活過程出現錯誤（可預期）：{e}")
        return False


def activate_excel_window(wb):
    """激活 Excel 視窗，使其成為前景視窗"""
    if not HAS_WIN32:
        print("提示：無法自動切換到 Excel 視窗（需要 pywin32 套件）")
        print("請手動點擊 Excel 視窗")
        return

    try:
        # 取得 Excel 視窗句柄
        excel_hwnd = wb.app.api.Hwnd

        # 檢查視窗是否存在
        if not win32gui.IsWindow(excel_hwnd):
            print("無法找到 Excel 視窗")
            return

        # 強制切換至前景（處理 Windows 前景視窗限制）
        if _force_foreground_window(excel_hwnd):
            print("✓ 已切換到 Excel 視窗")
        else:
            print("⚠️  視窗切換可能未完成，請手動點擊 Excel 視窗！")

    except Exception as e:
        logging.error(f"無法激活 Excel 視窗：{e}")


def activate_console_window(console_hwnd):
    """激活終端機視窗，使其成為前景視窗"""
    if not HAS_WIN32:
        print("提示：無法自動切換到終端機視窗（需要 pywin32 套件）")
        return

    try:
        # 嘗試找到正確的 Console 視窗
        current_hwnd = console_hwnd

        # 如果提供的句柄無效，嘗試搜尋 Console 視窗
        if not current_hwnd or not win32gui.IsWindow(current_hwnd):
            current_hwnd = _find_console_window()

        if current_hwnd and win32gui.IsWindow(current_hwnd):
            # 強制切換至前景（處理 Windows 前景視窗限制）
            if _force_foreground_window(current_hwnd):
                print("✓ 已切換到終端機視窗")
            else:
                print("⚠️  視窗切換可能未完成，請手動點擊終端機視窗！")
        else:
            print("提示：無法找到終端機視窗，請手動點擊終端機視窗")
    except Exception as e:
        print("提示：無法自動切換視窗，請手動點擊終端機視窗")
        logging.debug(f"SetForegroundWindow 失敗：{e}")


# =========================================================================
# Excel 導航輔助函數
# =========================================================================
# 因鍵盤監聽器以 suppress=True 攔截所有按鍵，Excel 無法收到方向鍵與翻頁鍵；
# 故由程式代收下列按鍵，改以 xlwings 對 Excel 執行等效操作：
#   Up/Down       : 作用儲存格上／下移一列
#   PgUp/PgDn     : 作用儲存格上／下移一個捲頁
#   Ctrl+PgUp/PgDn: 切換至前／後一個工作表
def _excel_move_active_cell(wb, row_delta: int):
    """將【作用儲存格】上下移動 row_delta 列（負值向上）"""
    try:
        xl = wb.app.api
        active_cell = xl.ActiveCell
        if active_cell is None:
            return
        new_row = max(1, active_cell.Row + row_delta)
        xl.ActiveSheet.Cells(new_row, active_cell.Column).Select()
        print(f"→ {xw.utils.col_name(active_cell.Column)}{new_row}")
    except Exception as e:
        logging.debug(f"移動作用儲存格失敗：{e}")


def _excel_page_scroll(wb, direction: int):
    """將【作用儲存格】上下移動一個捲頁（direction：+1 向下、-1 向上）"""
    try:
        xl = wb.app.api
        visible_rows = xl.ActiveWindow.VisibleRange.Rows.Count
        _excel_move_active_cell(wb, direction * visible_rows)
    except Exception as e:
        logging.debug(f"捲頁移動失敗：{e}")


def _excel_switch_sheet(wb, offset: int):
    """切換工作表（offset：+1 後一個、-1 前一個）"""
    try:
        idx = wb.sheets.active.index  # xlwings 之工作表索引為 1-based
        new_idx = idx + offset
        if 1 <= new_idx <= len(wb.sheets):
            new_sheet = wb.sheets[new_idx - 1]
            new_sheet.activate()
            print(f"→ 切換至工作表【{new_sheet.name}】")
        else:
            print(">> 已無前／後一個工作表可切換。")
    except Exception as e:
        logging.debug(f"切換工作表失敗：{e}")


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
            new_jin_kang_piau_im_ji_khoo_sheet=(args.new if hasattr(args, "new") else False),
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
        # 依【作用儲存格】所在之工作表，決定【資料來源工作表】（不再限定於單一工作表）
        source_sheet = wb.sheets.active
        source_sheet_name = source_sheet.name
        if source_sheet_name not in VALID_JI_KHOO_SHEETS:
            msg = f"作用工作表為【{source_sheet_name}】，本程式僅能於 {'、'.join(VALID_JI_KHOO_SHEETS)} 工作表使用，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_INVALID_INPUT

        # 取得【作用儲存格】
        active_cell_address = get_active_cell_address()
        active_cell = source_sheet.range(active_cell_address)
        row, col = excel_address_to_row_col(active_cell_address)
        # 自作用儲存格所在列，取得【漢字】（在A欄）
        han_ji = source_sheet.range((row, 1)).value
        tai_gi_im_piau = source_sheet.range((row, 2)).value

        # ----------------------------------------------------------------------
        # 依【作用儲存格】所在列（Row），自A欄取得【漢字】，B欄取得【台語音標】。
        # 並依上述資料判斷，是否程式還需繼續。
        # ----------------------------------------------------------------------
        if not is_han_ji(han_ji):
            msg = f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_SUCCESS
        elif source_sheet_name == KHUAT_JI_PIAU_SHEET and not tai_gi_im_piau == "N/A":
            # 僅【缺字表】需檢查：台語音標已有值（非 N/A）者，表示已查得讀音，跳過處理
            msg = f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，台語音標為【{tai_gi_im_piau}】，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_SUCCESS

        # ----------------------------------------------------------------------
        # 依取得之【漢字】，在漢字庫資料庫尋找【台語音標】
        # ----------------------------------------------------------------------
        cell = source_sheet.range((row, 1))
        tai_gi_im_piau = xls_cell._han_ji_ca_piau_im_kap_cu_tik(cell)

        # 使用者按 Enter 放棄選用（或查無讀音）：不作任何變更，正常結束本次作業
        if not tai_gi_im_piau:
            print(f">> 儲存格 {active_cell_address} 漢字【{han_ji}】：未選用讀音，不作任何變更。")
            return EXIT_CODE_SUCCESS

        # 依據使用者輸入之【台語音標】轉換為【漢字標音】
        han_ji_piau_im = xls_cell.convert_tai_gi_im_piau_to_han_ji_piau_im(
            tai_gi_im_piau=tai_gi_im_piau,
        )
        if not han_ji_piau_im:
            print(f"因異常狀況，終止作業......")
            print(f"{active_cell_address} 漢字：【{han_ji}】，台語音標：【{tai_gi_im_piau}】，漢字標音：【{han_ji_piau_im}】。")
            return EXIT_CODE_PROCESS_FAILURE

        # 將字典查得之【台語音標】回填【資料來源工作表】：
        # - 缺字表：回填 B 欄【台語音標】（原值為 N/A）
        # - 標音字庫/人工標音字庫：回填 C 欄【校正音標】
        if source_sheet_name == KHUAT_JI_PIAU_SHEET:
            source_sheet.range((row, 2)).value = tai_gi_im_piau
        else:
            source_sheet.range((row, 3)).value = tai_gi_im_piau

        # ----------------------------------------------------------------------
        # 將已查得之【台語音標】，及轉換所得之【漢字標音】，依據【缺字表】工作表在【座標】欄
        # 取得之【座標清單】，一一回填【漢字注音】工作表
        # ----------------------------------------------------------------------
        # 取得【座標】欄之【座標清單】：內容為字串，格式如 "(5, 17); (33, 8)"
        target_sheet = wb.sheets[program.hanji_piau_im_sheet_name]
        coordinates_str = source_sheet.range((row, 4)).value
        coordinate_tuples = []
        if coordinates_str:
            # 利用正規表達式，解析所有形如 (row, col) 之座標
            coordinate_tuples = convert_coord_list_str_to_tuples(
                coord_list=coordinates_str,
            )
            for tup in coordinate_tuples:
                coord_row, coord_col = int(tup[0]), int(tup[1])
                # 將【台語音標】和【漢字標音】寫入【漢字注音】工作表：
                # 【台語音標】在【漢字】儲存格上一列；【漢字標音】在下一列
                target_sheet.range((coord_row - 1, coord_col)).value = tai_gi_im_piau
                target_sheet.range((coord_row + 1, coord_col)).value = han_ji_piau_im

                # 顯示已完成的工作結果
                print(f"📌 ({coord_row}, {coord_col}) 漢字：{han_ji}，台語音標：{tai_gi_im_piau}")

        # -------------------------------------------------------------------------
        # 僅【缺字表】需執行：將原本在【缺字表】工作表之【資料紀錄】，寫入【標音字庫】
        # 工作表，以示該漢字已查得【台語音標】，並已在【漢字注音】工作表完成【漢字】的
        # 【台語音標】及【漢字標音】標注工作。
        # （【標音字庫】/【人工標音字庫】僅回填 C 欄【校正音標】，無需移轉紀錄。）
        # -------------------------------------------------------------------------
        if source_sheet_name == KHUAT_JI_PIAU_SHEET:
            xls_cell._update_piau_im_ji_khoo_from_khuat_ji_piau(
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                coordinates=coordinate_tuples,
            )

        # -------------------------------------------------------------------------
        # 更新資料庫中【漢字庫】資料表
        # -------------------------------------------------------------------------
        siong_iong_too_to_use = 0.8 if program.ue_im_lui_piat == "文讀音" else 0.6  # 根據語音類型設定常用度
        xls_cell.insert_or_update_to_db(
            table_name=program.table_name,
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            ue_im_lui_piat=program.ue_im_lui_piat,
            siong_iong_too=siong_iong_too_to_use,
        )
        # -------------------------------------------------------------------------
        # 因【標音字庫】依【漢字】之【座標】紀錄，更新【漢字注音】工作表中對映之【台語音標】及【漢字標音】；導致
        # 【作用儲存格】之 Excel Address 已變更，需將之校正回歸。
        # -------------------------------------------------------------------------
        source_sheet.activate()
        active_cell.select()  # 選取【作用儲存格】，以確保游標位置正確

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 處理作業結束
    # --------------------------------------------------------------------------
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 作業迴圈：鍵盤監聽模式（參考 a109 作法）
# =========================================================================
def run_keyboard_mode(wb, args) -> int:
    """
    鍵盤監聽模式：
    - <Enter> ：依【作用儲存格】執行【漢字查找讀音】作業
    - <Esc>   ：終止程式

    監聽器以 suppress=True 攔截按鍵，故在任何視窗按 <Enter> 均可觸發查音，
    且不會影響 Excel 儲存格內容（Excel 中按 Enter 不會使作用儲存格下移）。
    """
    # 監聽狀態（以 dict 供閉包函式讀寫）
    state = {"action": None, "running": True, "ctrl": False}

    CTRL_KEYS = (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, keyboard.Key.ctrl)

    def on_key_press(key):
        """鍵盤按下事件處理 - 只設置動作標記（動作於主執行緒執行）"""
        try:
            if key in CTRL_KEYS:
                state["ctrl"] = True
            elif key in (keyboard.Key.enter, keyboard.Key.space):
                # <Enter> 與 <Space> 均可觸發【漢字查找讀音】作業
                state["action"] = "process"
            elif key == keyboard.Key.up:
                state["action"] = "move_up"
            elif key == keyboard.Key.down:
                state["action"] = "move_down"
            elif key == keyboard.Key.page_up:
                # Ctrl+PgUp：切換至前一個工作表；PgUp：作用儲存格上移一個捲頁
                state["action"] = "prev_sheet" if state["ctrl"] else "page_up"
            elif key == keyboard.Key.page_down:
                # Ctrl+PgDn：切換至後一個工作表；PgDn：作用儲存格下移一個捲頁
                state["action"] = "next_sheet" if state["ctrl"] else "page_down"
            elif key == keyboard.Key.esc:
                state["action"] = "esc"
                state["running"] = False
                return False  # 停止監聽
        except Exception as e:
            logging.error(f"按鍵處理錯誤：{e}")

    def on_key_release(key):
        """鍵盤放開事件處理 - 追蹤 Ctrl 鍵狀態"""
        try:
            if key in CTRL_KEYS:
                state["ctrl"] = False
        except Exception as e:
            logging.error(f"按鍵處理錯誤：{e}")

    # 取得 Console 視窗句柄（程式啟動時，前景視窗即 Terminal）
    console_hwnd = None
    if HAS_WIN32:
        try:
            console_hwnd = win32gui.GetForegroundWindow()
        except Exception as e:
            logging.warning(f"無法取得 Console 視窗句柄：{e}")

    print("=" * 80)
    print("鍵盤監聽模式 - 操作說明：")
    print("  1. 以滑鼠或下列快捷鍵，於【缺字表／標音字庫／人工標音字庫】工作表，")
    print("     選定【作用儲存格】：")
    print("       ↑ / ↓        ：作用儲存格上／下移一列")
    print("       PgUp / PgDn  ：作用儲存格上／下移一個捲頁")
    print("       Ctrl+PgUp/PgDn：切換至前／後一個工作表")
    print("  2. 按 <Enter> 或 <Space> 鍵：執行【漢字查找讀音】作業（任一視窗皆可按）；")
    print("  3. 按 <Esc> 鍵：終止程式（並自動存檔）。")
    print("=" * 80)

    # 操作提示顯示完畢後，將視窗聚焦切回 Excel，方便使用者選取儲存格
    print("\n正在切換到 Excel 視窗...")
    activate_excel_window(wb)

    # 啟動鍵盤監聽（在背景執行緒，suppress=True 攔截按鍵，不讓 Excel 接收）
    listener = keyboard.Listener(on_press=on_key_press, on_release=on_key_release, suppress=True)
    listener.start()

    exit_code = EXIT_CODE_SUCCESS
    try:
        # 主迴圈：在主執行緒處理待執行的動作
        while state["running"]:
            action = state["action"]
            state["action"] = None

            if action == "process":
                # 暫停鍵盤監聽，使 process() 內之 input() 可正常接收使用者輸入
                listener.stop()
                time.sleep(0.3)

                # 切換到終端機視窗（確保使用者可以輸入）
                activate_console_window(console_hwnd)

                try:
                    exit_code = process(wb=wb, args=args)
                    if exit_code != EXIT_CODE_SUCCESS:
                        print(f"⚠️  處理結果：exit_code = {exit_code}")
                except Exception as e:
                    logging.error(f"處理錯誤：{e}")
                    print(f"❌ 錯誤：{e}")
                finally:
                    # 【漢字查找讀音】作業結束後，視窗聚焦切回 Excel，
                    # 方便使用者續選下一個【作用儲存格】
                    activate_excel_window(wb)

                    # 重新啟動鍵盤監聽
                    state["ctrl"] = False  # 重設 Ctrl 鍵狀態（監聽停止期間可能已放開）
                    listener = keyboard.Listener(on_press=on_key_press, on_release=on_key_release, suppress=True)
                    listener.start()
                    print("\n請選取下一個【作用儲存格】；按 <Enter> 或 <Space> 查音，按 <Esc> 結束。")

            elif action == "move_up":
                _excel_move_active_cell(wb, -1)
            elif action == "move_down":
                _excel_move_active_cell(wb, 1)
            elif action == "page_up":
                _excel_page_scroll(wb, -1)
            elif action == "page_down":
                _excel_page_scroll(wb, 1)
            elif action == "prev_sheet":
                _excel_switch_sheet(wb, -1)
            elif action == "next_sheet":
                _excel_switch_sheet(wb, 1)

            time.sleep(0.05)  # 避免 CPU 佔用過高

        print("\n按下 <Esc> 鍵，程式結束。")

    except KeyboardInterrupt:
        # Ctrl+C 亦可終止程式
        print("\n\n使用者中斷程式（Ctrl+C）")
    finally:
        if listener:
            listener.stop()

    return exit_code


# =========================================================================
# 作業迴圈：輸入模式（未安裝 pynput 時之備援）
# =========================================================================
def run_input_mode(wb, args) -> int:
    """輸入模式：於 Terminal 按 Enter 執行查音；按 Ctrl+C 終止程式"""
    print("=" * 80)
    print("輸入模式：請在【缺字表／標音字庫／人工標音字庫】任一工作表，")
    print("選擇儲存格後，回到 Terminal 按 Enter 查詢")
    print("按 Ctrl+C 終止程式（並自動存檔）")
    print("=" * 80)

    exit_code = EXIT_CODE_SUCCESS  # 避免使用者於首次處理前即按 Ctrl+C，致 exit_code 未定義
    while True:
        try:
            # 等待使用者按 Enter
            input("\n請在 Excel 選擇【作用儲存格】後按 Enter 繼續（Ctrl+C 終止）...")

            # 依【作用儲存格】所在之工作表執行處理（工作表判別於 process() 內進行）
            exit_code = process(wb=wb, args=args)
            if exit_code != EXIT_CODE_SUCCESS:
                print(f"⚠️  處理結果：exit_code = {exit_code}")

        except KeyboardInterrupt:
            # 無論 exit_code 為何，一律終止迴圈（修正原程式 Ctrl+C 無法終止之問題）
            print("\n\n使用者中斷程式（Ctrl+C）")
            print("=" * 70)
            break

        except Exception as e:
            logging.error(f"處理錯誤：{e}")
            print(f"❌ 錯誤：{e}")
            # 發生錯誤時繼續循環，不中斷程式
            continue

    return exit_code


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
        # 執行處理作業：依是否安裝 pynput 決定作業模式
        # ==================================================================
        if HAS_PYNPUT:
            exit_code = run_keyboard_mode(wb=wb, args=args)
        else:
            exit_code = run_input_mode(wb=wb, args=args)

        # ==================================================================
        # 程式終止前：儲存檔案
        # ==================================================================
        try:
            wb.save()
            file_path = wb.fullname
            logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案異常！", error=e)
            return EXIT_CODE_SAVE_FAILURE  # 作業異常終止：無法儲存檔案

        return exit_code

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
                    print(f"  台語音標：{item['台語音標']}, 常用度：{item.get('常用度', 'N/A')}, 說明：{item.get('摘要說明', 'N/A')}")
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
