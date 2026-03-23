"""
a260_依字典查得結果填入人工標音.py V0.0.4

在【漢字注音】工作表之【作用儲存格】，可以兩種方式輸入【人工標音】資料：
（1）自【自用字典】查得【台語音標】；（2）直接手動輸入【台語音標】/【台羅拼音】。

修改紀錄：
v0.0.1 2026-2-28: 初始版本，完成基本功能。
v0.0.2 2026-3-21: 修正查字典時，顯示所有讀音的預設值為 True。
v0.0.3 2026-3-22: 修正查字典後填入人工標音的邏輯，將【人工標音】、【台語音標】、【漢字標音】
    分別填入【作用儲存格】之上方一格、下方一格、同一格；並修正相關邏輯以確保資料正確填入。
v0.0.4 2026-3-23: 修正問題：當使用者放棄輸入【人工標音】時，即刻跳出 process 函式，避免後續
    更新【標音字庫】現有資料紀錄，引發錯誤。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from mod_excel_access import (
    excel_address_to_row_col,
    get_active_cell,
    get_active_cell_address,
    get_line_no_by_row,
    get_row_by_line_no,
)
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)

# 載入自訂模組/函式
from mod_標音 import is_han_ji, tlpa_tng_han_ji_piau_im
from mod_程式 import ExcelCell, Program

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


# =========================================================================
# 資料類別：儲存處理配置
# =========================================================================
class CellProcessor(ExcelCell):
    """
    個人字典查詢專用的儲存格處理器
    繼承自 ExcelCell
    覆蓋以下方法以實現個人字典查詢功能：
    - _process_cell(): 處理單一儲存格
    - _process_jin_kang_piau_im(): 處理人工標音邏輯
    其他方法繼承自父類別 ExcelCell
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

    def _za_ji_tain_au_thiam_jin_kang_piau_im(self, active_cell):
        """查字典後填入工標音"""
        piau_im_huat = self.program.piau_im_huat
        piau_im = self.program.piau_im
        tai_gi_im_piau = ""

        # 依據【作用儲存格】之【漢字】，從【自用字典】查詢【台語音標】
        tai_gi_im_piau = self._han_ji_ca_piau_im_kap_cu_tik(active_cell)
        if tai_gi_im_piau is None:
            return None, None

        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau,
        )

        active_cell.offset(-2, 0).value = tai_gi_im_piau  # 人工標音
        active_cell.offset(-1, 0).value = tai_gi_im_piau  # 台語音標
        active_cell.offset(1, 0).value = han_ji_piau_im  # 漢字標音

        return tai_gi_im_piau, han_ji_piau_im


# =============================================================================
# 作業主流程
# =============================================================================
def process(wb, args) -> int:
    """
    作業流程：
    1. 取得當前 Excel 作用儲存格 (漢字、座標)
    2. 計算【人工標音】位置與值
    3. 查詢【標音字庫】確認該座標是否已登錄
    4. 若【標正音標】為 'N/A'，則更新為【人工標音】

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
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        # --------------------------------------------------------------------------
        # 處理作業開始
        # --------------------------------------------------------------------------
        source_sheet_name = "漢字注音"

        # ----------------------------------------------------------------------
        # 取得【作用儲存格】
        # ----------------------------------------------------------------------
        # 指定【漢字注音】工作表為【作用工作表】
        source_sheet = wb.sheets[source_sheet_name]
        source_sheet.activate()

        active_cell_address = get_active_cell_address()
        active_cell = source_sheet.range(active_cell_address)
        row, col = excel_address_to_row_col(active_cell_address)
        current_line_no = get_line_no_by_row(current_row_no=row)  # 計算行號
        jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = (
            get_row_by_line_no(current_line_no)
        )
        source_sheet.range((han_ji_row, col)).select()  # 選取【漢字】儲存格，以確保游標位置正確
        source_sheet.activate()  # 重新激活工作表以刷新儲存格地址

        # 確認【作用儲存格】為【漢字】
        han_ji = source_sheet.range((han_ji_row, col)).value
        if not is_han_ji(han_ji):
            msg=f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
            print(f">> {msg}")
            return EXIT_CODE_SUCCESS

        # 確認【作用儲存格】的【漢字】有【台語音標】及【漢字標音】，否則可能是字典目前無此【漢字】之讀音資料，
        # 故後續之查字典作業應被略過，直接要求使用者輸入【台語音標】或【台羅拼音】。
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
        # 記錄原始的的【人工標音】
        original_jin_kang_piau_im = jin_kang_piau_im

        if not tai_gi_im_piau or not han_ji_piau_im:
            # ----------------------------------------------------------------------
            # 直接手動輸入人工標音，若是【作用儲存格】之【漢字】，可能字典尚未登錄此漢字之讀音資料
            # ----------------------------------------------------------------------
            msg = f"作用儲存格 {active_cell_address} 的漢字【{han_ji}】缺乏【台語音標】或【漢字標音】，可能是字典無此漢字之讀音資料，將略過查字典作業，直接要求使用者輸入【台語音標】或【台羅拼音】。"
            print(f">> {msg}")
            # 取得使用者輸入之【台語音標】或【台羅拼音】
            tai_gi_im_piau = xls_cell.get_user_input_piau_im(han_ji=han_ji)
            # 依據使用者輸入之【台語音標】轉換為【漢字標音】
            han_ji_piau_im = xls_cell._convert_tai_gi_im_piau_to_han_ji_piau_im(
                tai_gi_im_piau=tai_gi_im_piau,
            )

            source_sheet.range((tai_gi_im_piau_row, col)).value = tai_gi_im_piau
            source_sheet.range((han_ji_piau_im_row, col)).value = han_ji_piau_im
            source_sheet.range((jin_kang_piau_im_row, col)).value = jin_kang_piau_im
        else:
            # ----------------------------------------------------------------------
            # 查字典後填人工標音
            # ----------------------------------------------------------------------
            han_ji_position = (han_ji_row, col)
            print(
                f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}"
            )
            print(f"📌 漢字：{han_ji}")
            print(
                f"📌 人工標音：{jin_kang_piau_im}，台語音標：{tai_gi_im_piau}，漢字標音：{han_ji_piau_im}"
            )
            if not xls_cell._za_ji_tain_au_thiam_jin_kang_piau_im(
                active_cell=source_sheet.range((han_ji_row, col))
            ):
                return EXIT_CODE_SUCCESS    # 若使用者放棄變更，則結束作業流程

        # 透過【作用儲存格】取出處理後的【人工標音】、【台語音標】、【漢字標音】
        jin_kang_piau_im = source_sheet.range((jin_kang_piau_im_row, col)).value
        tai_gi_im_piau = source_sheet.range((tai_gi_im_piau_row, col)).value
        han_ji_piau_im = source_sheet.range((han_ji_piau_im_row, col)).value

        msg = f"【{han_ji}】變更為： [{jin_kang_piau_im}] / [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
        print(f">> 儲存格：{active_cell_address}，{msg}")

        # 將【台語音標】和【漢字標音】寫入【漢字注音】工作表之【作用儲存格】
        if not jin_kang_piau_im or original_jin_kang_piau_im == jin_kang_piau_im:    # 若【人工標音】儲存格未填入標音
            return EXIT_CODE_SUCCESS

        # -------------------------------------------------------------------------
        # 自【標音字庫】之【字庫表】(dict)，移除該【漢字】之記錄
        # -------------------------------------------------------------------------
        xls_cell._update_piau_im_ji_khoo_worksheet(cell=active_cell)
        # -------------------------------------------------------------------------
        # 在【人工標音字庫】之【字庫表】(dict)，新增該【漢字】之記錄
        # -------------------------------------------------------------------------
        xls_cell._update_jin_kang_piau_im_ji_khoo_worksheet(cell=active_cell)

        # -------------------------------------------------------------------------
        # 更新資料庫中【漢字庫】資料表
        # -------------------------------------------------------------------------
        siong_iong_too_to_use = (
            0.8 if program.ue_im_lui_piat == "文讀音" else 0.6
        )  # 根據語音類型設定常用度
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
        source_sheet.activate()  # 重新激活工作表以刷新儲存格地址
        active_cell.select()  # 選取【作用儲存格】，以確保游標位置正確

        logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
        # 再次拋出異常，讓外層函式能捕捉
        raise


# =========================================================================
# 主程式
# =========================================================================
def main(args) -> int:
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
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        exit_code = process(wb, args)
    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging_exc_error(msg=msg, error=None)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 要求畫面回到【漢字注音】工作表
        # wb.sheets['漢字注音'].activate()
        # 儲存檔案
        wb.save()
        file_path = wb.fullname
        logging_process_step(f"儲存檔案至路徑：{file_path}")

    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束作業
    # =========================================================================
    return EXIT_CODE_SUCCESS


def ut01():
    # 取得【作用中活頁簿】
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
    except Exception as e:
        msg = "無法找到作用中的 Excel 工作簿！"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_NO_FILE
    # 作業流程：獲取當前作用中的 Excel 儲存格
    sheet_name, cell_address = get_active_cell(wb)
    print(f"✅ 目前作用中的儲存格：{sheet_name} 工作表 -> {cell_address}")

    # 將 Excel 儲存格地址轉換為 (row, col) 格式
    row, col = excel_address_to_row_col(cell_address)
    print(f"📌 Excel 位址 {cell_address} 轉換為 (row, col): ({row}, {col})")

    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
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
    args = parser.parse_args()

    if args.test:
        # 執行測試
        ut01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code == EXIT_CODE_SUCCESS:
            print("程式正常完成！")
        else:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
