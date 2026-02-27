"""a215_漢字以人工標音處理作業.py v0.1

依據【作用儲存格】之【人工標音】欄位，處理人工手動標音作業。
 - 手動標音：在【人工標音】儲存格輸入完整的【台語音標】或【台羅拼音】（接受帶調符號的音標），
    手動標音會被記錄到【人工標音字庫】工作表；
 - 引用既有人工標音：輸入【=】符號，則【台語音標】將自【人工標音字庫】工作表查找；
 - 取消人工標音：輸入【-】符號，則取消人工標音，並從【人工標音字庫】工作表刪除該漢字的人工標音資料。

【處理規則】：
遇【人工標音】填入【引用既有的人工標音符號（=）】符號時，漢字的【台語音標】
自【人工標音字庫】工作表查找，並轉換成【漢字標音】。

若在【人工標音字庫】工作表找不到對映的【台語音標】，退而求其次，再自【標音字庫】
工作表查找；

若仍在【標音字庫】工作表亦找不到，則再退而求其次，自【字典】查找；如若仍找不到，
則將該漢字記錄到【缺字表】工作表。

更新紀錄：
v0.1.0 2026-02-27: 初始版本，實現基本功能：從【人工標音字庫】查找漢字的台語音標，並轉換成漢字標音；若找不到則記錄到【缺字表】。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import traceback
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from mod_excel_access import (
    get_active_cell_address,
)

# 載入自訂模組
from mod_logging import (
    init_logging,
    logging_exc_error,  # noqa: F401
    logging_exception,  # noqa: F401
    logging_process_step,  # noqa: F401
    logging_warning,  # noqa: F401
)
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
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        traceback.print_exc()
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 作業處理中
    # --------------------------------------------------------------------------
    try:
        # 處理工作表
        sheet_name = program.hanji_piau_im_sheet_name
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 取得作用儲存格
        active_cell_address = get_active_cell_address()
        active_cell = sheet.range(active_cell_address)
        # row, col = excel_address_to_row_col(active_cell_address)
        # current_line_no = get_line_no_by_row(current_row_no=row)  # 計算行號
        # jin_kang_piau_im_row, tai_gi_im_piau_row, han_ji_row, han_ji_piau_im_row = (
        #     get_row_by_line_no(current_line_no)
        # )
        xls_cell._process_jin_kang_piau_im(cell=active_cell)

    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        traceback.print_exc()
        return EXIT_CODE_PROCESS_FAILURE

    # --------------------------------------------------------------------------
    # 處理作業結束
    # --------------------------------------------------------------------------
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
            traceback.print_exc()
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
        try:
            wb.save()
            file_path = wb.fullname
            logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案異常！", error=e)
            return EXIT_CODE_SAVE_FAILURE  # 作業異當終止：無法儲存檔案
        return EXIT_CODE_SUCCESS

    # =========================================================================
    # 結束程式
    # =========================================================================
    print("\n")
    print("=" * 80)
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    if exit_code == EXIT_CODE_SUCCESS:
        return EXIT_CODE_SUCCESS  # 作業正常結束
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
                    print(
                        f"  台語音標：{item['台語音標']}, 常用度：{item.get('常用度', 'N/A')}, 說明：{item.get('摘要說明', 'N/A')}"
                    )
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
