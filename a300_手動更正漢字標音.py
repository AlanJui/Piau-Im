# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw

from a330_以標音字庫更新漢字注音工作表 import process_update_hanji_zu_im_sheet_by_piau_im_ji_khoo
from mod_excel_access import excel_address_to_row_col, get_active_cell
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,  # noqa: F401
)

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
init_logging()


# =============================================================================
# 作業主流程
# =============================================================================
def _show_separtor_line(source_sheet_name: str, target_sheet_name: str):
    print('\n\n')
    print("=" * 100)
    print(f"使用【{source_sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
    print("=" * 100)

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
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        # 設定作業程序名稱
        procedure_name = "初始化作業程序"

        # 建立程式物件
        program = Program(wb, args, hanji_piau_im_sheet_name='漢字注音')

        # 建立儲存格處理器
        xls_cell = None
        if args.new:
            # 建立【標音字庫工作表】
            xls_cell = ExcelCell(
                program=program,
                new_jin_kang_piau_im_ji_khoo_sheet=True,
                new_piau_im_ji_khoo_sheet=True,
                new_khuat_ji_piau_sheet=True,
            )
        else:
            # xls_cell = ExcelCell(program=program)
            xls_cell = ExcelCell(
                program=program,
                new_jin_kang_piau_im_ji_khoo_sheet=False,
                new_piau_im_ji_khoo_sheet=False,
                new_khuat_ji_piau_sheet=False,
            )

    except Exception as e:
        logging_exception(msg=f"{procedure_name}，發生作業異常，終止處理！", error=e)
        raise

    #--------------------------------------------------------------------------
    # 處理作業開始
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 開始處理流程作業！==========>")

    try:
        # 取得目標工作表
        target_sheet_name = '漢字注音'
        sheet_name = target_sheet_name
        han_ji_piau_im_sheet = wb.sheets[sheet_name]
        han_ji_piau_im_sheet.activate()
        logging_process_step("已完成作業所需之初始化設定！")
    except Exception as e:
        logging_exc_error(msg=f"找不到【{sheet_name}】工作表 ！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #-----------------------------------------------------------------------------
    # 將【缺字表】工作表，已填入【台語音標】之資料，登錄至【標音字庫】工作表
    # 使用【缺字表】工作表中的【校正音標】，更正【漢字注音】工作表中之【台語音標】、【漢字標音】；
    # 並依【缺字表】工作表中的【台語音標】儲存格內容，更新【標音字庫】工作表中之【台語音標】及【校正音標】欄位
    #-----------------------------------------------------------------------------
    try:
        source_sheet_name = '缺字表'
        sheet_name = source_sheet_name
        print('\n\n')
        print("=" * 100)
        # print(f"使用【{sheet_name}】工作表的【校正音標】欄位，更新【{target_sheet_name}】工作表之【台語音標】、【漢字標音】：")
        logging_process_step(
            msg=f"以【{source_sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print("=" * 100)
        # 將【缺字表】工作表中的【台語音標】儲存格內容，更新至【標音字庫】工作表中之【台語音標】及【校正音標】欄位
        # update_khuat_ji_piau(wb=wb)
        # 依據【缺字表】工作表紀錄，並參考【漢字注音】工作表在【人工標音】欄位的內容，更新【缺字表】工作表中的【校正音標】及【台語音標】欄位
        # 即使用者為【漢字】補入查找不到的【台語音標】時，若是在【缺字表】工作表中之【校正音標】直接填寫
        # 則應執行 a310*.py 程式；但使用者若是在【漢字注音】工作表中之【人工標音】欄位填寫，則應執行 a320*.py 程式
        # a300*.py 之本程式
        xls_cell.update_hanji_zu_im_sheet_by_khuat_ji_piau(
            source_sheet_name=source_sheet_name,
            target_sheet_name=target_sheet_name,
        )
        # 將所有【標音字庫工作表】對映之字典物件，回存 Excel 活頁簿檔案(Workbook)
        xls_cell.save_all_piau_im_ji_khoo_dicts()
    except Exception as e:
        logging_exc_error(msg=f"處理【{sheet_name}】作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #-----------------------------------------------------------------------------
    # 將【漢字注音】工作表，【漢字】填入【人工標音】內容，登錄至【人工標音字庫】及
    # 【標音字庫】工作表
    #-----------------------------------------------------------------------------
    try:
        # 使用【漢字注音】工作表作為【目標工作表】
        target_sheet_name = '漢字注音'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")

        # 使用【人工標音字庫】，作為來源工作表中的【校正音標】欄位，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
        source_sheet_name = '人工標音字庫'
        sheet_name = source_sheet_name
        _show_separtor_line(source_sheet_name=source_sheet_name, target_sheet_name=target_sheet_name)
        xls_cell.update_hanji_zu_im_sheet_by_jin_kang_piau_im_ji_khoo(
            source_sheet_name=source_sheet_name,
            target_sheet_name=target_sheet_name,
        )
        # 使用【缺字表】，作為來源工作表中的【校正音標】欄位，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
        source_sheet_name = '缺字表'
        sheet_name = source_sheet_name
        _show_separtor_line(source_sheet_name=source_sheet_name, target_sheet_name=target_sheet_name)
        xls_cell.update_hanji_zu_im_sheet_by_khuat_ji_piau(
            source_sheet_name=source_sheet_name,
            target_sheet_name=target_sheet_name,
        )
        # 使用【標音字庫】工作表中的【校正音標】欄位，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
        # 執行 a200_查找及填入漢字標音：可在【漢字注音】工作表，直接標注【人工標音】者，並更新【人工標音字庫】工作表
        source_sheet_name = '標音字庫'
        sheet_name = source_sheet_name
        _show_separtor_line(source_sheet_name=source_sheet_name, target_sheet_name=target_sheet_name)
        xls_cell.update_hanji_zu_im_sheet_by_piau_im_ji_khoo(
            source_sheet_name=source_sheet_name,
            target_sheet_name=target_sheet_name,
        )
    except Exception as e:
        logging_exc_error(msg=f"使用【{sheet_name}】工作表，更新【漢字注音】工作表，發生作業異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #-----------------------------------------------------------------------------
    # 掃瞄【漢字注音】工作表，對於【儲存格】填寫【人工標音】、【引用人工標音】或
    # 【去除人工標音】等特殊狀況之【漢字】，更新【人工標音工作表】、【標音字庫工作表】內容。
    #-----------------------------------------------------------------------------

    #-----------------------------------------------------------------------------
    # 根據【標音字庫】工作表，更新【漢字注音】工作表中的【台語音標】及【漢字標音】欄位
    #-----------------------------------------------------------------------------
    try:
        sheet_name = '標音字庫'
        logging_process_step(f"以【{sheet_name}】工作表之【校正音標】，更新【{target_sheet_name}】工作表之【台語音標】與【漢字標音】！")
        print('\n\n')
        print("=" * 100)
        print(f"使用【{sheet_name}】工作表中的【校正音標】，更新【漢字注音】工作表中的【台語音標】：")
        print("=" * 100)
        process_update_hanji_zu_im_sheet_by_piau_im_ji_khoo(wb=wb)
    except Exception as e:
        logging_exc_error(msg=f"處理以【{sheet_name}】更新【漢字注音】工作表之作業，發生執行異常！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    #------------------------------------------------------------------------------
    # 處理作業結束
    #------------------------------------------------------------------------------
    han_ji_piau_im_sheet.activate()

    print('=' * 80)
    logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
    logging_process_step("<=========== 完成處理流程作業！==========>")

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
        logging.error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 儲存檔案
        if not Program.save_workbook_as_new_file(wb=wb):
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

    # =========================================================================
    # (5) 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS


def ut01():
    # 取得【作用中活頁簿】
    wb = None
    # 取得【作用中活頁簿】
    try:
        wb = xw.apps.active.books.active    # 取得 Excel 作用中的活頁簿檔案
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
        ut01()
    else:
        # 從 Excel 呼叫
        exit_code = main(args)
        if exit_code == EXIT_CODE_SUCCESS:
            print("程式正常完成！")
        else:
            print(f"程式異常終止，錯誤代碼為: {exit_code}")
            sys.exit(exit_code)
