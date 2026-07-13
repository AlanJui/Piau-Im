"""
a215_漢字以人工標音處理作業.py v0.1.1

依【作用儲存格】所在處，指定【漢字標音】工作表裡的【漢字】，其讀音採用【人工標音】。

X 依據【作用儲存格】之【人工標音】欄位，處理人工手動標音作業。
X  - 手動標音：在【人工標音】儲存格輸入完整的【台語音標】或【台羅拼音】（接受帶調符號的音標），
X     手動標音會被記錄到【人工標音字庫】工作表；
X  - 引用既有人工標音：輸入【=】符號，則【台語音標】將自【人工標音字庫】工作表查找；
X  - 取消人工標音：輸入【-】符號，則取消人工標音，並從【人工標音字庫】工作表刪除該漢字的人工標音資料。
X
X 【處理規則】：
X 遇【人工標音】填入【引用既有的人工標音符號（=）】符號時，漢字的【台語音標】
X 自【人工標音字庫】工作表查找，並轉換成【漢字標音】。
X
X 若在【人工標音字庫】工作表找不到對映的【台語音標】，退而求其次，再自【標音字庫】
X 工作表查找；
X
X 若仍在【標音字庫】工作表亦找不到，則再退而求其次，自【字典】查找；如若仍找不到，
X 則將該漢字記錄到【缺字表】工作表。

更新紀錄：
v0.1.0 2026-02-27: 初始版本，實現基本功能：從【人工標音字庫】查找漢字的台語音標，並轉換成漢字標音；若找不到則記錄到【缺字表】。
v0.1.1 2026-02-28: 修正【作用儲存格】處理邏輯，可以自動矯正【漢字】儲存格之座標，避免使用者操作不慎之問題。
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
from mod_標音 import is_han_ji
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

    def _process_sheet(self):
        """處理【漢字注音】工作表之作用儲存格訂正／人工標音作業。"""
        try:
            # --------------------------------------------------------------------------
            # 取得【作用儲存格】
            # --------------------------------------------------------------------------
            sheet_name = self.program.hanji_piau_im_sheet_name
            source_sheet = self.program.wb.sheets[sheet_name]

            # 取得【漢字標音】工作表的【作用儲存格】
            han_ji_cell = self.get_han_ji_cell_with_active_cell()
            # 自【漢字】儲存格取得【位址】（Excel Address）及【座標】 (row, col)
            active_cell_address = han_ji_cell.address.replace("$", "")
            han_ji_row, col = han_ji_cell.row, han_ji_cell.column

            # 確認【作用儲存格】為【漢字】
            han_ji = han_ji_cell.offset(0, 0).value
            if not is_han_ji(han_ji):
                msg = f"作用儲存格 {active_cell_address} 的漢字為【{han_ji}】，屬於標點符號或特殊符號，跳過處理。"
                print(f">> {msg}")
                return EXIT_CODE_SUCCESS

            # 確認【作用儲存格】的【漢字】有【台語音標】及【漢字標音】，否則可能是字典目前無此【漢字】之讀音資料，
            # 故後續之查字典作業應被略過，直接要求使用者輸入【台語音標】或【台羅拼音】。
            jin_kang_piau_im = han_ji_cell.offset(-2, 0).value
            tai_gi_im_piau = han_ji_cell.offset(-1, 0).value
            han_ji_piau_im = han_ji_cell.offset(1, 0).value
            new_jin_kang_piau_im = None

            if not tai_gi_im_piau or not han_ji_piau_im:
                # ----------------------------------------------------------------------
                # 直接手動輸入人工標音，若是【作用儲存格】之【漢字】，可能字典尚未登錄此漢字之讀音資料
                # ----------------------------------------------------------------------
                msg = (
                    f"作用儲存格 {active_cell_address} 的漢字【{han_ji}】缺乏【台語音標】或【漢字標音】，"
                    f"後續作業無法進行，請先手動補全【台語音標】與【漢字標音】，再執行本程式。"
                )
                print(f">> {msg}")
                return EXIT_CODE_PROCESS_FAILURE
            elif not jin_kang_piau_im:
                msg = (
                    f"作用儲存格 {active_cell_address} 的漢字【{han_ji}】未填【人工標音】，"
                    f"請手動輸入完整的【台語音標】或【台羅拼音】（接受帶調符號的音標），"
                    f"或輸入【=】符號以引用既有人工標音，或輸入【-】符號以取消人工標音。"
                )
                print(f">> {msg}")
                return EXIT_CODE_PROCESS_FAILURE

            # ----------------------------------------------------------------------
            # 查字典後填人工標音
            # ----------------------------------------------------------------------
            han_ji_position = (han_ji_row, col)
            print(f"📌 作用儲存格：{active_cell_address} ==> 漢字儲存格座標：{han_ji_position}")
            print(f"📌 漢字：{han_ji}")
            print(f"📌 人工標音：{jin_kang_piau_im}，原台語音標：{tai_gi_im_piau}，原漢字標音：{han_ji_piau_im}")

            # 依據【作用儲存格】輸入之【人工標音】，轉換【漢字標音】
            tai_gi_im_piau = jin_kang_piau_im
            han_ji_piau_im = self._convert_tai_gi_im_piau_to_han_ji_piau_im(
                tai_gi_im_piau=tai_gi_im_piau,
            )

            # 若是【沒有查到漢字之台語音標】或是【使用者終止手動輸入】，則程式至此終止。
            if not jin_kang_piau_im or not tai_gi_im_piau and not han_ji_piau_im:
                return EXIT_CODE_PROCESS_FAILURE
            # 將查尋/輸入取得之【台語音標】視作【人工標音】
            new_jin_kang_piau_im = tai_gi_im_piau if tai_gi_im_piau else None
            # 在 Console 回報目前作業狀態
            msg = f"【{han_ji}】變更為： [{jin_kang_piau_im}] / [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
            print(f">> 儲存格：{active_cell_address}，{msg}")

            # -------------------------------------------------------------------------
            # 原先由程式自動標注【台語音標】的【漢字】，改成【人工標音】。
            # -------------------------------------------------------------------------
            self._change_han_ji_from_tai_gi_im_piau_to_jin_kang_piau_im(
                row=han_ji_row,
                col=col,
                han_ji=han_ji,
                jin_kang_piau_im=new_jin_kang_piau_im,
                tai_gi_im_piau=tai_gi_im_piau,
                han_ji_piau_im=han_ji_piau_im,
            )

            # -------------------------------------------------------------------------
            # 更新資料庫中【漢字庫】資料表
            # -------------------------------------------------------------------------
            siong_iong_too_to_use = 0.8 if self.program.ue_im_lui_piat == "文讀音" else 0.6  # 根據語音類型設定常用度
            self.insert_or_update_to_db(
                table_name=self.program.table_name,
                han_ji=han_ji,
                tai_gi_im_piau=tai_gi_im_piau,
                ue_im_lui_piat=self.program.ue_im_lui_piat,
                siong_iong_too=siong_iong_too_to_use,
            )
            # -------------------------------------------------------------------------
            # 因【標音字庫】依【漢字】之【座標】紀錄，更新【漢字注音】工作表中對映之【台語音標】及【漢字標音】；導致
            # 【作用儲存格】之 Excel Address 已變更，需將之校正回歸。
            # -------------------------------------------------------------------------
            source_sheet.activate()  # 重新激活工作表以刷新儲存格地址
            han_ji_cell.select()  # 選取【作用儲存格】，以確保游標位置正確

            logging_process_step(msg="已完成【台語音標】和【漢字標音】標注工作。")
            return EXIT_CODE_SUCCESS
        except Exception as e:
            logging_exception(msg="自動為【漢字】查找【台語音標】作業，發生例外！", error=e)
            raise


# =========================================================================
# 主要處理函數
# =========================================================================
def process(wb, args) -> int:
    """
    作業流程：
    1. 初始化 Program 與 CellProcessor
    2. 呼叫 CellProcessor._process_sheet() 處理作用儲存格

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    logging_process_step("<=========== 作業開始！==========>")
    # --------------------------------------------------------------------------
    # 作業初始化
    # --------------------------------------------------------------------------
    try:
        program = Program(wb=wb, args=args, hanji_piau_im_sheet_name="漢字注音")

        # 建立儲存格處理器（繼承自 ExcelCell）
        xls_cell = CellProcessor(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 作業處理中
    # --------------------------------------------------------------------------
    try:
        xls_cell._process_sheet()
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


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
    # 取得【作用中活頁簿】
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
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS  # 作業正常結束


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
