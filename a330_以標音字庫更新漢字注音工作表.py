# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

# 載入自訂模組/函式
from mod_excel_access import save_as_new_file

# 載入自訂模組/函式
from mod_程式 import ExcelCell, Program

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
from mod_logging import (  # noqa: E402
    init_logging,
    logging_exc_error,
    logging_exception,  # noqa: F401
    logging_process_step,
)

init_logging()


# =========================================================================
# 本程式主要處理作業程序
# =========================================================================
def process(wb, args) -> int:
    """
    更新【漢字注音】表中【台語音標】儲存格的內容，依據【標音字庫】中的【校正音標】欄位進行更新，並將【校正音標】覆蓋至原【台語音標】。

    Args:
        wb: Excel Workbook 物件
        args: 命令列參數

    Returns:
        處理結果代碼
    """
    #--------------------------------------------------------------------------
    # 作業開始
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    #--------------------------------------------------------------------------
    # 作業初始化
    #--------------------------------------------------------------------------
    try:
        program = Program(wb, args, hanji_piau_im_sheet_name='漢字注音')

        # 建立儲存格處理器
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )
    except Exception as e:
        logging_exc_error(msg="處理作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #--------------------------------------------------------------------------
    # 處理作業開始
    #--------------------------------------------------------------------------
    source_sheet_name="標音字庫"
    target_sheet_name="漢字注音"
    msg = f'使用【{source_sheet_name}】工作表，更新【{target_sheet_name}】工作表......'
    print('\n')
    print("=" * 80)
    logging_process_step(msg)

    try:
        sheet_name = source_sheet_name
        wb.sheets[sheet_name].activate()
        exit_code = xls_cell.update_hanji_zu_im_sheet_by_piau_im_ji_khoo(
            source_sheet_name=source_sheet_name,
            target_sheet_name=target_sheet_name,
        )
    except Exception as e:
        logging_exc_error(msg=f"處理【{sheet_name}】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        return exit_code

    #-------------------------------------------------------------------------
    # 更新資料庫 & 關閉資料庫連線
    #-------------------------------------------------------------------------
    # 關閉資料庫連線
    if xls_cell.db_manager:
        xls_cell.db_manager.disconnect()
        logging_process_step("已關閉資料庫連線")

    #--------------------------------------------------------------------------
    # 作業結束
    #--------------------------------------------------------------------------
    print('\n')
    logging_process_step("<=========== 作業結束！==========>")
    return EXIT_CODE_SUCCESS


# =========================================================================
# 程式主要作業流程
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
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
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
        logging_exception(msg="作業異常終止！", error=e)
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


# =============================================================================
# 測試程式
# =============================================================================
def test_01() -> int:
    """
    測試程式主要作業流程
    """
    print("\n\n")
    print("=" * 100)
    print("執行測試：test_01()")
    print("=" * 100)
    # 執行主要作業流程
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
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
        exit_code = test_01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)

    # 只在命令列執行時使用 sys.exit()，避免在調試環境中引發 SystemExit 例外
    if exit_code != EXIT_CODE_SUCCESS:
        sys.exit(exit_code)

