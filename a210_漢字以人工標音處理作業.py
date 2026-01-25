"""
    簡單說明作業流程如下：
    遇【作用儲存格】填入【引用既有的漢字標音】符號（【=】）時，漢字的【台語音標】
    自【人工標音字庫】工作表查找，並轉換成【漢字標音】。

    在【漢字注音】工作表，若使用者曾對某漢字以【人工標音】儲存格手動標音過，則再
    次遇到相同之漢字時，若在【人工標音】儲存格填入【=】符號（表示引用既有的標音），
    則使用者可省去重新標音的麻煩；而程式會負責自【人工標音字庫】工作表查找該漢字的
    【台語音標】，並轉換成【漢字標音】填入對應的儲存格。

    顧及使用者可能會有記憶錯誤的狀況發生，若在【人工標音字庫】工作表找不到對應的
    【台語音標】時，程式會再自【標音字庫】工作表查找一次，若仍找不到，則將該漢字
    記錄到【缺字表】工作表，以便後續處理。
"""
# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
from pathlib import Path
from typing import Tuple

# 載入第三方套件
import xlwings as xw

# 載入自訂模組
from mod_file_access import save_as_new_file
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
    logging_exception,  # noqa: F401
    logging_process_step,  # noqa: F401
    logging_warning,  # noqa: F401
)
from mod_帶調符音標 import is_han_ji
from mod_標音 import ca_ji_kiat_ko_tng_piau_im
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
# 資料類別：儲存處理配置
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

    def _resolve_manual_annotation(self, han_ji: str, jin_kang_val: any) -> str | None:
        """
        解析人工標音內容，處理特殊符號 (=, #) 與一般標音。

        依據 PRG-a210 規則：
        1. 若內容為 '#'：取消人工指定，回傳 None (後續將回歸資料庫查找)。
        2. 若內容為 '='：自【人工標音字庫】查找該漢字的音標。
        3. 若為其他內容：直接回傳該內容作為音標。
        4. 若為空：回傳 None。

        Args:
            han_ji (str): 當前處理的漢字。
            jin_kang_val (any): 人工標音儲存格的原始值。

        Returns:
            str | None: 決定的台語音標。若為 None，表示需進行資料庫查找。
        """
        # 轉為字串並去除前後空白，若為 None 則變為空字串
        jin_kang_str = str(jin_kang_val).strip() if jin_kang_val is not None else ""

        # Case 0: 無內容 -> 回歸資料庫查找
        if not jin_kang_str:
            return None

        # Case 1: 內容為 '#' -> 強制取消人工指定，回歸資料庫查找
        if jin_kang_str == '#':
            logging.debug(f"漢字 '{han_ji}' 人工標音為 '#'，強制回歸資料庫查找。")
            return None

        # Case 2: 內容為 '=' -> 引用人工標音工作表 (從已載入的字庫中查找)
        if jin_kang_str == '=':
            # 檢查字庫中是否有此漢字
            # jin_kang_ji_khoo 預期是 JiKhooDict 物件，或 dict 結構
            if han_ji in self.jin_kang_ji_khoo:
                # 取得該漢字的所有標音紀錄
                # 結構預期為: {漢字: {音標: [次數, 座標列表], ...}}
                piau_im_variants = self.jin_kang_ji_khoo[han_ji]

                if piau_im_variants:
                    # 策略：取用第一個找到的音標 (或可依需求改為取用頻率最高的)
                    # 這裡實作：取字典的第一個 Key
                    target_piau_im = next(iter(piau_im_variants))
                    logging.info(f"漢字 '{han_ji}' 引用人工標音 '='，查得: {target_piau_im}")
                    return target_piau_im
                else:
                    logging.warning(f"漢字 '{han_ji}' 設為 '='，但在人工標音字庫中無音標資料，將回歸資料庫查找。")
                    return None
            else:
                logging.warning(f"漢字 '{han_ji}' 設為 '='，但在人工標音字庫中查無此字，將回歸資料庫查找。")
                return None

        # Case 3: 一般內容 -> 直接使用該內容作為音標
        return jin_kang_str

    def _convert_piau_im(self, result: list) -> Tuple[str, str]:
        """
        將查詢結果轉換為音標

        Args:
            result: 查詢結果列表

        Returns:
            (tai_gi_im_piau, han_ji_piau_im)
        """
        # 使用原有的轉換邏輯
        # 這裡需要適配 result 的格式
        # 假設 result 是從 HanJiSuTian 回傳的格式
        tai_gi_im_piau, han_ji_piau_im = ca_ji_kiat_ko_tng_piau_im(
            result=result,
            han_ji_khoo=self.program.han_ji_khoo,
            piau_im=self.program.piau_im,
            piau_im_huat=self.program.piau_im_huat
        )
        return tai_gi_im_piau, han_ji_piau_im


    def _process_han_ji(
        self,
        han_ji: str,
        cell,
        row: int,
        col: int,
    ) -> Tuple[str, bool]:
        """處理漢字"""
        if han_ji == '':
            return "【空白】", False

        # 使用 HanJiTian 查詢漢字讀音
        result = self.program.ji_tian.han_ji_ca_piau_im(
            han_ji=han_ji,
            ue_im_lui_piat=self.program.ue_im_lui_piat
        )

        # 查無此字
        if not result:
            self.program.khuat_ji_piau_ji_khoo_dict.add_or_update_entry(
                han_ji=han_ji,
                tai_gi_im_piau='',
                hau_ziann_im_piau='N/A',
                coordinates=(row, col)
            )
            return f"【{han_ji}】查無此字！", False

        # 轉換音標
        tai_gi_im_piau, han_ji_piau_im = self._convert_piau_im(result)

        # 寫入儲存格
        cell.offset(-1, 0).value = tai_gi_im_piau  # 上方儲存格：台語音標
        cell.offset(1, 0).value = han_ji_piau_im    # 下方儲存格：漢字標音

        # 記錄到標音字庫
        self.program.piau_im_ji_khoo_dict.add_or_update_entry(
            han_ji=han_ji,
            tai_gi_im_piau=tai_gi_im_piau,
            hau_ziann_im_piau='N/A',
            coordinates=(row, col)
        )

        return f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】", False


    def _process_jin_kang_piau_im(self, cell, row, col):
        """
        處理單一漢字儲存格的主邏輯
        """
        # 1. 取得漢字與人工標音儲存格內容
        han_ji = cell.value
        # 人工標音位於漢字上方兩列 (Row - 2)
        jin_kang_cell = cell.offset(-2, 0)
        jin_kang_val = jin_kang_cell.value

        # 2. 解析人工標音 (處理 =, # 邏輯)
        manual_tai_gi_im = self._resolve_manual_annotation(han_ji, jin_kang_val)

        if manual_tai_gi_im:
            # === A. 使用人工標音 ===
            tai_gi_im_piau = manual_tai_gi_im

            # 呼叫標音轉換模組，產生對應的漢字標音 (如：台羅、方音等)
            # 假設 self.config.piau_im 是 PiauIm 物件
            han_ji_piau_im = self.program.piau_im.han_ji_piau_im_tng_huan(
                self.program.piau_im_huat,
                self.program.piau_im.split_tai_gi_im_piau(tai_gi_im_piau)
            )

            # 寫回 Excel (台語音標欄位 row-1, 漢字標音欄位 row+1)
            cell.offset(-1, 0).value = tai_gi_im_piau
            cell.offset(1, 0).value = han_ji_piau_im

            # 記錄到人工標音字庫 (更新引用計數或座標)
            # self.jin_kang_ji_khoo.add_entry(...)

        else:
            # === B. 回歸資料庫查找 ===
            # 當人工標音為空、為 '#' 或 '=' 查找失敗時執行
            # self._process_han_ji_from_db(cell, han_ji)
            self._process_han_ji(
                han_ji=han_ji,
                cell=cell,
                row=row,
                col=col,
            )

    def _process_han_ji_from_db(self, cell, han_ji):
        # 實作原本的資料庫查找邏輯...
        pass

    def _show_msg(self, row: int, col: int, msg: str):
        """顯示處理訊息"""
        # 顯示處理進度
        col_name = xw.utils.col_name(col)
        print(f"【{col_name}{row}】({row}, {col}) = {msg}")

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
        self._reset_cell_style(cell)

        # 取得【漢字】儲存格內容
        cell_value = cell.value

        # 檢查是否有【人工標音】
        jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
        if jin_kang_piau_im and str(jin_kang_piau_im).strip() != "":
            self._process_jin_kang_piau_im(cell, row, col)

        # 依據【漢字】儲存格內容進行處理
        if cell_value == 'φ':
            self._show_msg(row, col, "【文字終結】")
            return  1   # 文章終結符號
        elif cell_value == '\n':
            self._show_msg(row, col, "【換行】")
            return  2   #【換行】
        elif not is_han_ji(cell_value):
            self._process_non_han_ji(cell_value)
            return 3    # 標點符號或空白
        else:
            self._process_han_ji(cell_value, cell, row, col)
            return  0  # 漢字

    def _process_sheet(self, sheet):
        """處理整個工作表"""
        EOF = False # 是否到達文件結尾
        line = 1

        config = self.program
        for row in range(config.start_row, config.end_row, config.ROWS_PER_LINE):
            status_code = 0
            EOL = False # 是否到達行尾
            # 設定作用儲存格為列首
            sheet.range((row, 1)).select()

            # 逐欄處理
            for col in range(config.start_col, config.end_col):
                cell = sheet.range((row, col))
                # 設定作用儲存格為目前儲存格
                cell.select()

                # 處理儲存格
                # status_code:
                # 0 = 儲存格內容為：漢字
                # 1 = 儲存格內容為：文字終結符號
                # 2 = 儲存格內容為：換行符號
                # 3 = 儲存格內容為：空白、標點符號等非漢字字元
                status_code = self._process_cell(cell, row, col)

                # 檢查是否終結、跳出列處理迴圈
                if status_code == 1:
                    EOF = True
                    break
                elif status_code == 2:
                    EOL = True
                    break

            # 檢查是否到達結尾
            if EOF or EOL or line > config.TOTAL_LINES:
                break

            # 換行顯示
            if col == config.end_col - 1:
                print('\n')

            line += 1


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
    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        #--------------------------------------------------------------------------
        # 初始化 Program 配置
        #--------------------------------------------------------------------------
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name='漢字注音')

        # 建立萌典專用的儲存格處理器（繼承自 ExcelCell）
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=args.new if hasattr(args, 'new') else False,
            new_piau_im_ji_khoo_sheet=args.new if hasattr(args, 'new') else False,
            new_khuat_ji_piau_sheet=args.new if hasattr(args, 'new') else False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 作業處理中
    #--------------------------------------------------------------------------
    try:
        # 處理工作表
        sheet_name = program.hanji_piau_im_sheet_name
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        xls_cell._process_sheet(sheet=sheet)

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 處理作業結束
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業結束！==========>")
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
    # 取得【作用中活頁簿】
    wb = None
    try:
        # 嘗試從 Excel 呼叫取得（RunPython）
        wb = xw.Book.caller()
    except Exception:
        # 若失敗，則取得作用中的活頁簿
        try:
            wb = xw.apps.active.books.active
        except Exception as e:
            logging_exc_error(msg="無法找到作用中的 Excel 工作簿！", error=e)
            return EXIT_CODE_NO_FILE

    if not wb:
        logging_exc_error(msg="無法取得 Excel 活頁簿！", error=None)
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    exit_code = process(wb, args)

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    if exit_code == EXIT_CODE_SUCCESS:
        file_path = save_as_new_file(wb=wb)
        if not file_path:
            logging_exc_error(msg="儲存檔案失敗！", error=None)
            exit_code = EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
        else:
            logging_process_step(f"儲存檔案至路徑：{file_path}")

    # =========================================================================
    # 結束程式
    # =========================================================================
    print('\n')
    print('=' * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    if exit_code == EXIT_CODE_SUCCESS:
        return EXIT_CODE_SUCCESS    # 作業正常結束
    else:
        msg = f"程式異常終止，返回失敗碼：{exit_code}"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

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
    DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')

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
        description='缺字表修正後續作業程式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a000.py          # 執行一般模式
  python a000.py -new     # 建立新的字庫工作表
  python a000.py -test    # 執行測試模式
'''
        )
    parser.add_argument(
        '--test',
        action='store_true',
        help='執行測試模式',
    )
    parser.add_argument(
        '--new',
        action='store_true',
        help='建立新的標音字庫工作表',
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