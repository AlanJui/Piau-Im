"""
程式名稱：a320_以字庫工作表更新漢字庫資料庫.py v0.0.1

功能：
透過每個已完成編輯、校正的【漢字標音】Excel 檔，更新資料庫【漢字庫】資料表中【最近揀用時間】欄位
的資料，以便提高【漢字】的某【讀音】具有較高之權重，在【漢字查找讀音】作業，可被優先揀用。

使用的工作表為：【標音字庫】與【人工標音字庫】：

1. 以【標音字庫】工作表之【漢字】、【台語音標】欄位，更新資料庫【漢字庫】資料表；
2. 以【人工標音字庫】工作表之【漢字】、【台語音標】欄位，更新資料庫【漢字庫】資料表。

更新紀錄：
 - v0.0.1 2026-07-15: 新增功能。
"""

# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

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
    依序以【標音字庫】、【人工標音字庫】工作表之【漢字】（A欄）、【台語音標】（B欄），
    更新資料庫【漢字庫】資料表：
    - 紀錄（漢字＋台羅音標）已存在：依【語音類型】更新【常用度】（文讀音 0.8／白話音 0.6），
      並更新【更新時間】與【最近揀用時間】；
    - 紀錄不存在：新增一筆，【常用度】依【語音類型】設定，並填上【最近揀用時間】。

    註：【人工標音字庫】置於【標音字庫】之後處理。因【人工標音】為使用者明確指定之讀音，
    其【最近揀用時間】較晚，於【漢字查找讀音】排序時可獲得最高優先權。

    本程式不更動活頁簿中任何工作表之內容。

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
    # 處理作業開始：依序以各來源工作表，更新資料庫【漢字庫】資料表
    # 註：【人工標音字庫】置於【標音字庫】之後處理，因【人工標音】為使用者
    #     明確指定之讀音，其【最近揀用時間】較晚，於查音排序時可獲得最高優先權。
    #--------------------------------------------------------------------------
    source_sheet_names = ["標音字庫", "人工標音字庫"]

    for sheet_name in source_sheet_names:
        msg = f'使用【{sheet_name}】工作表，更新資料庫【漢字庫】資料表......'
        print('\n')
        print("=" * 80)
        logging_process_step(msg)

        try:
            if sheet_name not in [sheet.name for sheet in wb.sheets]:
                logging_process_step(f"活頁簿中無【{sheet_name}】工作表，略過此步驟。")
                continue
            wb.sheets[sheet_name].activate()
            exit_code = xls_cell.update_han_ji_khoo_db_by_sheet(sheet_name=sheet_name)
            if exit_code != EXIT_CODE_SUCCESS:
                return exit_code
        except ValueError as e:
            # 工作表內無資料：非屬異常，略過此步驟即可
            logging_process_step(f"【{sheet_name}】工作表內無資料，略過此步驟。（{e}）")
        except Exception as e:
            logging_exc_error(msg=f"處理【{sheet_name}】作業異常！", error=e)
            return EXIT_CODE_PROCESS_FAILURE

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
        msg = f"作業程序發生異常，終止執行：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_PROCESS_FAILURE

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"處理作業發生異常，終止程式執行：{program_name}（處理作業程序，返回失敗碼）"
        logging_exc_error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 本程式僅更新資料庫，不更動活頁簿內容，故無需儲存 Excel 檔案。
    # =========================================================================

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
        description='以【標音字庫】與【人工標音字庫】工作表，更新資料庫【漢字庫】資料表',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用範例：
  python a350_以字庫工作表更新漢字庫.py          # 執行一般模式
  python a350_以字庫工作表更新漢字庫.py --test   # 執行測試模式
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

