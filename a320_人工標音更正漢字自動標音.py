# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import argparse
import logging
import os
import re
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import save_as_new_file
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)
from mod_帶調符音標 import kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho
from mod_標音 import (
    tlpa_tng_han_ji_piau_im,  # 台語音標轉台語音標
)
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
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
init_logging()

# =========================================================================
# 程式區域函式
# =========================================================================
# def insert_or_update_to_db(db_path, table_name: str, han_ji: str, tai_gi_im_piau: str, piau_im_huat: str):
#     """
#     將【漢字】與【台語音標】插入或更新至資料庫。

#     :param db_path: 資料庫檔案路徑。
#     :param table_name: 資料表名稱。
#     :param han_ji: 漢字。
#     :param tai_gi_im_piau: 台語音標。
#     """
#     conn = sqlite3.connect(db_path)
#     cursor = conn.cursor()

#     # 確保資料表存在
#     cursor.execute(f"""
#     CREATE TABLE IF NOT EXISTS {table_name} (
#         識別號 INTEGER NOT NULL UNIQUE PRIMARY KEY AUTOINCREMENT,
#         漢字 TEXT,
#         台羅音標 TEXT,
#         常用度 REAL,
#         摘要說明 TEXT,
#         建立時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime')),
#         更新時間 TEXT NOT NULL DEFAULT (DATETIME('now', 'localtime'))
#     );
#     """)

#     # 檢查是否已存在該漢字
#     cursor.execute(f"SELECT 識別號 FROM {table_name} WHERE 漢字 = ?", (han_ji,))
#     row = cursor.fetchone()

#     siong_iong_too = 0.8 if piau_im_huat == "文讀音" else 0.6
#     if row:
#         # 更新資料
#         cursor.execute(f"""
#         UPDATE {table_name}
#         SET 台羅音標 = ?, 更新時間 = ?
#         WHERE 識別號 = ?;
#         """, (tai_gi_im_piau, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row[0]))
#     else:
#         # 若語音類型為：【文讀音】，設定【常用度】欄位值為 0.8
#         cursor.execute(f"""
#         INSERT INTO {table_name} (漢字, 台羅音標, 常用度, 摘要說明)
#         VALUES (?, ?, ?, NULL);
#         """, (han_ji, tai_gi_im_piau, siong_iong_too))

#     conn.commit()
#     conn.close()

#--------------------------------------------------------------------------
# 重整【標音字庫】查詢表：重整【標音字庫】工作表使用之 Dict
# 依據【缺字表】工作表之【漢字】+【台語音標】資料，在【標音字庫】工作表【添增】此筆資料紀錄
#--------------------------------------------------------------------------
# def tiau_zing_piau_im_ji_khoo_dict(
#     piau_im_ji_khoo_dict:JiKhooDict,
#     han_ji:str,
#     tai_gi_im_piau:str,
#     row:int, col:int
# ) -> bool:
#     # Step 1: 在【標音字庫】搜尋該筆【漢字】+【台語音標】
#     existing_entries = piau_im_ji_khoo_dict.ji_khoo_dict.get(han_ji, [])

#     # 標記是否找到
#     entry_found = False

#     for existing_entry in existing_entries:
#         # Step 2: 若找到，移除該筆資料內的座標
#         if (row, col) in existing_entry["coordinates"]:
#             existing_entry["coordinates"].remove((row, col))
#         entry_found = True
#         break  # 找到即可離開迴圈

#     # Step 3: 將此筆資料（校正音標為 'N/A'）於【標音字庫】底端新增
#     piau_im_ji_khoo_dict.add_entry(
#         han_ji=han_ji,
#         tai_gi_im_piau=tai_gi_im_piau,
#         kenn_ziann_im_piau="N/A",  # 預設值
#         coordinates=(row, col)
#     )
#     return entry_found

def update_han_ji_zu_im_piau_by_jin_kang_piau_im_ji_khoo_piau(
    program:Program,
    xls_cell:ExcelCell,
    source_sheet_name:str='人工標音字庫',
    target_sheet_name:str='漢字注音',
):
    """
    讀取 Excel 檔案，依據【人工標音字庫】工作表中的資料執行下列作業：
      1. 由 A 欄讀取漢字，從 B 欄取得原始輸入之【台語音標】，並轉換為 TLPA+ 格式，然後更新 C 欄（校正音標）。
      2. 從 D 欄讀取座標字串（可能含有多組座標），每組座標指向【漢字注音】工作表中該漢字儲存格，
         將【缺字表】取得之【台語音標】，填入【漢字注音】工作表之【台語音標】欄位（於【漢字】儲存格上方一列（row - 1））;
         並在【漢字】儲存格下方一列（row + 1）填入【漢字標音】。
    """
    # 取得本函式所需之【選項】參數
    wb = program.wb
    piau_im_huat = program.piau_im_huat
    piau_im = program.piau_im
    try:
        # 取得【來源工作表】（人工標音字庫）
        source_sheet = wb.sheets[source_sheet_name]
        # 取得【目標工作表】（漢字注音）
        target_sheet = wb.sheets[target_sheet_name]
        # 建立【標音字庫】查詢表（dict）
        piau_im_ji_khoo_dict  = xls_cell.piau_im_ji_khoo_dict
    except Exception as e:
        logging_exc_error(f"找不到名為『{source_sheet_name}』的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    #-------------------------------------------------------------------------
    # 在【人工標音字庫】工作表中，逐列讀取資料進行處理：【校正音標】欄（C）有填音標者，
    # 將【校正音標】正規化為 TLPA+ 格式，並更新【台語音標】欄（B）；
    # 並依據【座標】欄（D）內容，將【校正音標】填入【漢字注音】工作表中相對應之【台語音標】儲存格，
    # 以及使用【校正音標】轉換後之【漢字標音】填入【漢字注音】工作表中相對應之【漢字標音】儲存格。
    #-------------------------------------------------------------------------
    row = 2  # 從第 2 列開始（跳過標題列）
    while True:
        han_ji = source_sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束迴圈
            break

        # 取得原始【台語音標】並轉換為 TLPA+ 格式
        org_tai_gi_im_piau = source_sheet.range(f"B{row}").value
        if org_tai_gi_im_piau == "N/A" or not org_tai_gi_im_piau:  # 若【台語音標】欄為空，則結束迴圈
            row += 1
            continue
        if org_tai_gi_im_piau and kam_si_u_tiau_hu(org_tai_gi_im_piau):
            tlpa_im_piau = tng_im_piau(org_tai_gi_im_piau)   # 將【音標】使用之【拼音字母】轉換成【TLPA拼音字母】；【音標調符】仍保持
            tlpa_im_piau = tng_tiau_ho(tlpa_im_piau).lower()  # 將【音標調符】轉換成【數值調號】
        else:
            tlpa_im_piau = org_tai_gi_im_piau  # 若非帶調符音標，則直接使用原音標

        # 讀取【缺字表】中【座標】欄（D 欄）的內容
        # 欄中內容可能含有多組座標，如 "(5, 17); (33, 8); (77, 5)"，表【漢字注音】工作表中有多處需要更新
        hau_ziann_im_piau = tlpa_im_piau  # 預設【校正音標】為 TLPA+ 格式
        coordinates_str = source_sheet.range(f"D{row}").value
        print(f"{row-1}. (A{row}) 【{han_ji}】==> {coordinates_str} ： 台語音標：{org_tai_gi_im_piau}, 校正音標：{hau_ziann_im_piau}\n")

        # 將【座標】欄位內容解析成 (row, col) 座標：此座標指向【漢字注音】工作表中之【漢字】儲存格位置
        # tai_gi_im_piau = tlpa_im_piau
        tai_gi_im_piau = hau_ziann_im_piau  # 使用【校正音標】填入【漢字注音】工作表之【台語音標】欄位
        if coordinates_str:
            # 利用正規表達式解析所有形如 (row, col) 的座標
            coordinate_tuples = re.findall(r"\((\d+)\s*,\s*(\d+)\)", coordinates_str)
            for tup in coordinate_tuples:
                try:
                    r_coord = int(tup[0])
                    c_coord = int(tup[1])
                except ValueError:
                    continue  # 若轉換失敗，跳過該組座標

                han_ji_cell = (r_coord, c_coord)  # 漢字儲存格座標

                # 根據說明，【台語音標】應填入漢字儲存格上方一列 (row - 1)，相同欄位
                target_row = r_coord - 1
                tai_gi_im_piau_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                target_sheet.range(tai_gi_im_piau_cell).value = tai_gi_im_piau
                excel_address = target_sheet.range(tai_gi_im_piau_cell).address
                excel_address = excel_address.replace("$", "")  # 去除 "$" 符號
                print(f"   台語音標：【{tai_gi_im_piau}】，填入【漢字注音】工作表之 {excel_address} 儲存格 = {tai_gi_im_piau_cell}")

                # 轉換【台語音標】，取得【漢字標音】
                han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                    piau_im=piau_im, piau_im_huat=piau_im_huat, tai_gi_im_piau=tai_gi_im_piau
                )

                # 將【漢字標音】填入【漢字注音】工作表，【漢字】儲存格下之【漢字標音】儲存格（即：row + 1)
                target_row = r_coord + 1
                han_ji_piau_im_cell = (target_row, c_coord)

                # 將【校正音標】填入【漢字注音】工作表漢字之【台語音標】儲存格
                target_sheet.range(han_ji_piau_im_cell).value = han_ji_piau_im
                excel_address = target_sheet.range(han_ji_piau_im_cell).address
                excel_address = excel_address.replace("$", "")  # 去除 "$" 符號
                print(f"   漢字標音：【{han_ji_piau_im}】，填入【漢字注音】工作表之 {excel_address} 儲存格 = {han_ji_piau_im_cell}\n")

                # 將【漢字注音】工作表之【漢字】儲存格之底色，重置為【無底色】
                target_sheet.range(han_ji_cell).color = None

                # 更新【標音字庫】工作表之資料紀錄
                xls_cell.tiau_zing_piau_im_ji_khoo_dict(
                    piau_im_ji_khoo_dict=piau_im_ji_khoo_dict,
                    han_ji=han_ji,
                    tai_gi_im_piau=tai_gi_im_piau,
                    row=r_coord,
                    col=c_coord,
                )

        row += 1  # 讀取下一列

    # 依據 Dict 內容，更新【標音字庫】工作表之資料紀錄
    piau_im_ji_khoo_dict.write_to_excel_sheet(wb)

    return EXIT_CODE_SUCCESS


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
    # 作業初始化
    #--------------------------------------------------------------------------
    logging_process_step("<=========== 作業開始！==========>")

    try:
        #--------------------------------------------------------------------------
        # 初始化 process config
        #--------------------------------------------------------------------------
        program = Program(wb, args, hanji_piau_im_sheet='漢字注音')

        # 建立儲存格處理器
        # xls_cell = ExcelCell(program=program)
        xls_cell = ExcelCell(
            program=program,
            new_jin_kang_piau_im_ji_khoo_sheet=False,
            new_piau_im_ji_khoo_sheet=False,
            new_khuat_ji_piau_sheet=False,
        )

        #--------------------------------------------------------------------------
        # 處理作業開始
        #--------------------------------------------------------------------------
        # 處理工作表
        sheet_name = '漢字注音'
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # ------------------------------------------------------------------------------
        # 以【人工標音字庫】工作表中各【校正音標】欄之注音，更新【漢字注音】工作表
        # 中【台語音標】及【漢字標音】儲存格內容
        # ------------------------------------------------------------------------------
        try:
            # 以【人工標音字庫】工作表中各【校正音標】欄之注音，更新【漢字注音】工作表
            update_han_ji_zu_im_piau_by_jin_kang_piau_im_ji_khoo_piau(
                program=program,
                xls_cell=xls_cell,
                source_sheet_name="人工標音字庫",
                target_sheet_name="漢字注音",
            )
        except Exception as e:
            logging_exc_error(msg=f"處理【{sheet_name}】作業異常！", error=e)
            return EXIT_CODE_PROCESS_FAILURE

        print("------------------------------------------------------")
        msg = f'使用【{sheet_name}】工作表，更新【漢字注音】工作表之己完成！'
        logging_process_step(msg)

        #-------------------------------------------------------------------------
        # 將【缺字表】之【漢字】與【台語音標】存入【漢字庫】作業
        #-------------------------------------------------------------------------
        try:
            sheet_name = '人工標音字庫'
            wb.sheets[sheet_name].activate()
            xls_cell.update_han_ji_khoo_db_by_sheet(sheet_name=sheet_name)
        except Exception as e:
            logging_exc_error(
                msg=f"將【{sheet_name}】之【漢字】與【台語音標】存入【漢字庫】作業，發生執行異常！",
                error=e)
            return EXIT_CODE_PROCESS_FAILURE

        #--------------------------------------------------------------------------
        # 處理作業結束
        #--------------------------------------------------------------------------
        # 寫回字庫到 Excel
        xls_cell.save_all_piau_im_ji_khoo_dict()

        print('\n')
        logging_process_step("<=========== 作業結束！==========>")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        msg=f"處理作業，發生異常！ ==> error = {e}"
        logging.exception(msg)
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
        msg = f"程式異常終止：{program_name}"
        logging_exception(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    if exit_code != EXIT_CODE_SUCCESS:
        msg = f"程式異常終止：{program_name}（非例外，而是返回失敗碼）"
        logging.error(msg)
        return EXIT_CODE_PROCESS_FAILURE

    # =========================================================================
    # (4) 儲存檔案
    # =========================================================================
    try:
        # 要求畫面回到【漢字注音】工作表
        # wb.sheets['漢字注音'].activate()
        # 儲存檔案
        file_path = save_as_new_file(wb=wb)
        if not file_path:
            logging_exc_error(msg="儲存檔案失敗！", error=None)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
        else:
            logging_process_step(f"儲存檔案至路徑：{file_path}")
    except Exception as e:
        logging_exception(msg="儲存檔案失敗！", error=e)
        return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


# =============================================================================
# 測試程式
# =============================================================================
def test_01():
    """
    測試程式主要作業流程
    """
    print("\n\n")
    print("=" * 100)
    print("執行測試：test_01()")
    # 執行主要作業流程
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式作業模式切換
# =============================================================================
if __name__ == "__main__":
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
        sys.exit(test_01())
    else:
        # 從 Excel 呼叫
        sys.exit(main())