# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from pathlib import Path

# 載入第三方套件
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組
from mod_excel_access import delete_sheet_by_name, get_value_by_name
from mod_file_access import load_module_function, save_as_new_file
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_標音 import ca_ji_kiat_ko_tng_piau_im  # 查字典得台語音標及漢字標音
from mod_標音 import convert_tl_with_tiau_hu_to_tlpa  # 去除台語音標的聲調符號
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import tlpa_tng_han_ji_piau_im  # 漢字標音物件
from mod_標音 import PiauIm

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
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()


# =========================================================================
# 程式區域函式
# =========================================================================
def jin_kang_piau_im_cu_han_ji_piau_im(wb, jin_kang_piau_im: str, piau_im: PiauIm, piau_im_huat: str):
    """
    人工標音取【台語音標】
    """

    if '〔' in jin_kang_piau_im and '〕' in jin_kang_piau_im:
        # 將人工輸入的〔台語音標〕轉換成【方音符號】
        im_piau = jin_kang_piau_im.split('〔')[1].split('〕')[0]
        tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(im_piau)
        # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau
        )
    elif '【' in jin_kang_piau_im and '】' in jin_kang_piau_im:
        # 將人工輸入的【方音符號】轉換成【台語音標】
        han_ji_piau_im = jin_kang_piau_im.split('【')[1].split('】')[0]
        siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
        # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        tai_gi_im_piau = piau_im.hong_im_tng_tai_gi_im_piau(
            siann=siann,
            un=un,
            tiau=tiau)['台語音標']
    else:
        # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
        tai_gi_im_piau = convert_tl_with_tiau_hu_to_tlpa(jin_kang_piau_im)
        # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
        han_ji_piau_im = tlpa_tng_han_ji_piau_im(
            piau_im=piau_im,
            piau_im_huat=piau_im_huat,
            tai_gi_im_piau=tai_gi_im_piau
        )

    return tai_gi_im_piau, han_ji_piau_im


def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im', new_jin_kang_piau_im__piau:bool=False):
    """查漢字讀音：依【漢字】查找【台語音標】，並依指定之【標音方法】輸出【漢字標音】"""
    try:
        # 載入【漢字庫】查找函數
        han_ji_ca_piau_im = load_module_function(module_name, function_name)

        # 連接指定資料庫
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()

        # 建置 PiauIm 物件，供作漢字拼音轉換作業
        han_ji_khoo_field = '漢字庫'
        han_ji_khoo_name = get_value_by_name(wb=wb, name=han_ji_khoo_field)
        piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)            # 指定漢字自動查找使用的【漢字庫】
        piau_im_huat = get_value_by_name(wb=wb, name='標音方法')    # 指定【台語音標】轉換成【漢字標音】的方法

        # 建置自動及人工漢字標音字庫工作表：（1）【標音字庫】；（2）【人工標音字】；（3）【缺字表】
        khuat_ji_piau_name = '缺字表'
        delete_sheet_by_name(wb=wb, sheet_name=khuat_ji_piau_name)
        khuat_ji_piau_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
                                    wb=wb,
                                    sheet_name=khuat_ji_piau_name)

        piau_im_sheet_name = '標音字庫'
        delete_sheet_by_name(wb=wb, sheet_name=piau_im_sheet_name)
        piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
                                wb=wb,
                                sheet_name=piau_im_sheet_name)

        jin_kang_piau_im_sheet_name='人工標音字庫'
        if new_jin_kang_piau_im__piau:
            delete_sheet_by_name(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)
        jin_kang_piau_im_ji_khoo = JiKhooDict.create_ji_khoo_dict_from_sheet(
            wb=wb,
            sheet_name=jin_kang_piau_im_sheet_name)

        # 指定【漢字注音】工作表為【作用工作表】
        sheet = wb.sheets[sheet_name]
        sheet.activate()

        # 設定起始及結束的【列】位址（【第5列】、【第9列】、【第13列】等列）
        TOTAL_LINES = int(wb.names['每頁總列數'].refers_to_range.value)
        ROWS_PER_LINE = 4
        start_row = 5
        end_row = start_row + (TOTAL_LINES * ROWS_PER_LINE)

        # 設定起始及結束的【欄】位址（【D欄=4】到【R欄=18】）
        CHARS_PER_ROW = int(wb.names['每列總字數'].refers_to_range.value)
        start_col = 4
        end_col = start_col + CHARS_PER_ROW

        # 逐列處理作業
        EOF = False
        line = 1
        for row in range(start_row, end_row, ROWS_PER_LINE):
            # 設定【作用儲存格】為列首
            Two_Empty_Cells = 0
            sheet.range((row, 1)).select()

            # 逐欄取出漢字處理
            for col in range(start_col, end_col):
                # Initialize variables to avoid using them before assignment
                tai_gi_im_piau = ""
                han_ji_piau_im = ""

                # 取得當前儲存格內含值
                # han_ji_u_piau_im = False
                msg = ""
                cell = sheet.range((row, col))
                # 將文字顏色設為【自動】（黑色）
                cell.font.color = (0, 0, 0)  # 設定為黑色
                # 將儲存格的填滿色彩設為【無填滿】
                cell.color = None

                cell_value = cell.value
                jin_kang_piau_im = cell.offset(-2, 0).value  # 人工標音
                if cell_value == 'φ':
                    EOF = True
                    msg = "【文字終結】"
                elif cell_value == '\n':
                    msg = "【換行】"
                elif cell_value == None or cell_value.strip() == "":  # 若儲存格內無值
                    if Two_Empty_Cells == 0:
                        Two_Empty_Cells += 1
                    elif Two_Empty_Cells == 1:
                        EOF = True
                    msg = "【空缺】"    # 表【儲存格】未填入任何字/符，不同於【空白】字元
                # 若不為【標點符號】，則以【漢字】處理
                elif is_punctuation(cell_value):
                    msg = f"{cell_value}【標點符號】"
                else:
                    # 查找漢字讀音
                    han_ji = cell_value
                    if jin_kang_piau_im and han_ji != '':
                        tai_gi_im_piau, han_ji_piau_im = jin_kang_piau_im_cu_han_ji_piau_im(
                            wb=wb,
                            jin_kang_piau_im=jin_kang_piau_im,
                            piau_im=piau_im,
                            piau_im_huat=piau_im_huat)
                        # 將【台語音標】和【漢字標音】寫入儲存格
                        sheet.range((row - 1, col)).value = tai_gi_im_piau
                        sheet.range((row + 1, col)).value = han_ji_piau_im
                        msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】《人工標音》]"
                        # 【標音字庫】添加或更新【漢字】資料
                        # jin_kang_piau_im_ji_khoo.add_or_update_entry(
                        jin_kang_piau_im_ji_khoo.add_entry(
                            han_ji=han_ji,
                            tai_gi_im_piau=tai_gi_im_piau,
                            kenn_ziann_im_piau=jin_kang_piau_im,
                            coordinates=(row, col)
                        )
                    else:
                        # 自【漢字庫】查找作業
                        result = han_ji_ca_piau_im(cursor=cursor,
                                                    han_ji=han_ji,
                                                    ue_im_lui_piat=ue_im_lui_piat)
                        # 若【漢字庫】查無此字，登錄至【缺字表】
                        if not result:
                            # khuat_ji_piau_ji_khoo.add_or_update_entry(
                            khuat_ji_piau_ji_khoo.add_entry(
                                han_ji=han_ji,
                                tai_gi_im_piau='',
                                kenn_ziann_im_piau='N/A',
                                coordinates=(row, col)
                            )
                            msg = f"【{han_ji}】查無此字！"
                        else:
                            # 依【漢字庫】查找結果，輸出【台語音標】和【漢字標音】
                            tai_gi_im_piau, han_ji_piau_im = ca_ji_kiat_ko_tng_piau_im(
                                result=result,
                                han_ji_khoo=han_ji_khoo,
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat
                            )
                            # 將【台語音標】和【漢字標音】寫入儲存格
                            sheet.range((row - 1, col)).value = tai_gi_im_piau
                            sheet.range((row + 1, col)).value = han_ji_piau_im
                            msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"
                            # 【標音字庫】添加或更新【漢字】資料
                            # piau_im_ji_khoo.add_or_update_entry(
                            piau_im_ji_khoo.add_entry(
                                han_ji=han_ji,
                                tai_gi_im_piau=tai_gi_im_piau,
                                kenn_ziann_im_piau='N/A',
                                coordinates=(row, col)
                            )
                # 顯示處理進度
                col_name = xw.utils.col_name(col)   # 取得欄位名稱
                print(f"【{xw.utils.col_name(col)}{row}】({row}, {col_name}) = {msg}")

                # 若讀到【換行】或【文字終結】，跳出逐欄取字迴圈
                if msg == "【換行】" or EOF:
                    break

            # -----------------------------------------------------------------
            # 若已到【結尾】或【超過總行數】，則跳出迴圈
            if EOF or line > TOTAL_LINES:
                # 將【標音字庫】、【人工標音字庫】、【缺字表】三個字典，寫入 Excel 工作表
                khuat_ji_piau_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=khuat_ji_piau_name)
                piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
                jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)
                break
            # 每當處理一行 15 個漢字後，亦換到下一行
            if col == end_col - 1: print('\n')
            line += 1

        #----------------------------------------------------------------------
        # 作業結束前處理
        #----------------------------------------------------------------------
        # 關閉資料庫連線
        conn.close()
        logging_process_step("已完成【台語音標】和【漢字標音】標注工作。")
        return EXIT_CODE_SUCCESS
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging.exception("自動為【漢字】查找【台語音標】作業，發生例外！")
        # 再次拋出異常，讓外層函式能捕捉
        raise


def process(wb):
    logging_process_step("<----------- 作業開始！---------->")
    # ---------------------------------------------------------------------
    # 重設【漢字】儲存格文字及底色格式
    # ---------------------------------------------------------------------
    # reset_han_ji_cells(wb=wb)

    # ------------------------------------------------------------------------------
    # 為漢字查找讀音，漢字上方填：【台語音標】；漢字下方填使用者指定之【漢字標音】
    # ------------------------------------------------------------------------------
    han_ji_khoo_field = '漢字庫'
    han_ji_khoo_name = get_value_by_name(wb=wb, name=han_ji_khoo_field) # 取得【漢字庫】名稱：河洛話、廣韻
    ue_im_lui_piat = get_value_by_name(wb, '語音類型')  # 取得【語音類型】，判別使用【白話音】或【文讀音】何者。
    db_name = 'Ho_Lok_Ue.db' if han_ji_khoo_name == '河洛話' else 'Kong_Un.db'

    if han_ji_khoo_name == '河洛話':
        module_name = 'mod_河洛話'
    else:
        module_name = 'mod_廣韻'
    function_name = 'han_ji_ca_piau_im'

    try:
        if han_ji_khoo_name == "河洛話" and ue_im_lui_piat == "白話音":
            ca_han_ji_thak_im(
                wb=wb,
                sheet_name="漢字注音",
                cell="V3",
                ue_im_lui_piat=ue_im_lui_piat,  # "白話音"
                han_ji_khoo=han_ji_khoo_name,   # "河洛話",
                db_name=db_name,                # "Ho_Lok_Ue.db",
                module_name=module_name,        # "mod_河洛話",
                function_name=function_name,    # "han_ji_ca_piau_im",
                new_jin_kang_piau_im__piau=True # 新建人工標音字庫工作表
            )
        elif han_ji_khoo_name == "河洛話" and ue_im_lui_piat == "文讀音":
            ca_han_ji_thak_im(
                wb=wb,
                sheet_name="漢字注音",
                cell="V3",
                ue_im_lui_piat=ue_im_lui_piat,  # "文讀音"
                han_ji_khoo=han_ji_khoo_name,   # "河洛話",
                db_name=db_name,                # "Ho_Lok_Ue.db",
                module_name=module_name,        # "mod_河洛話",
                function_name=function_name,    # "han_ji_ca_piau_im",
                new_jin_kang_piau_im__piau=True # 新建人工標音字庫工作表
            )
        elif han_ji_khoo_name == "廣韻":
            ca_han_ji_thak_im(
                wb=wb,
                sheet_name="漢字注音",
                cell="V3",
                ue_im_lui_piat="文讀音",
                han_ji_khoo="廣韻",
                db_name="Kong_Un.db",
                module_name="mod_廣韻",
                function_name="han_ji_ca_piau_im",
            )
        else:
            msg = "無法執行漢字標音作業，請確認【env】工作表【語音類型】及【漢字庫】欄位的設定是否正確！"
            logging_exc_error(msg=msg, error=None)
            return EXIT_CODE_INVALID_INPUT
    except Exception as e:
        logging_exc_error(msg="在查找漢字標音時發生錯誤！", error=e)
        return EXIT_CODE_PROCESS_FAILURE

    print("------------------------------------------------------")
    msg = f'自動為【漢字】查找【台語音標】作業己完成！'
    logging_process_step(msg)
    logging_process_step(f'【語音類型】：{ue_im_lui_piat}')
    logging_process_step(f'【漢字庫】：{db_name}')

    #--------------------------------------------------------------------------
    # 結束作業
    #--------------------------------------------------------------------------
    # 要求畫面回到【漢字注音】工作表
    wb.sheets['漢字注音'].activate()
    # 作業正常結束
    logging_process_step("<----------- 作業結束！---------->")
    return EXIT_CODE_SUCCESS

# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # (0) 程式初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
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
        print(f"發生錯誤: {e}")
        logging.error(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE

    # 若無法取得【作用中活頁簿】，則因無法繼續作業，故返回【作業異常終止代碼】結束。
    if not wb:
        return EXIT_CODE_NO_FILE

    # =========================================================================
    # (3) 執行【處理作業】
    # =========================================================================
    try:
        status_code = process(wb)
        if status_code != EXIT_CODE_SUCCESS:
            msg = f"程式異常終止：{program_name}"
            logging_exc_error(msg=msg, error=e)
            return EXIT_CODE_PROCESS_FAILURE

    except Exception as e:
        msg = f"程式異常終止：{program_name}"
        logging_exc_error(msg=msg, error=e)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        #--------------------------------------------------------------------------
        # 儲存檔案
        #--------------------------------------------------------------------------
        try:
            # 要求畫面回到【漢字注音】工作表
            wb.sheets['漢字注音'].activate()
            # 儲存檔案
            file_path = save_as_new_file(wb=wb)
            if not file_path:
                logging_exc_error(msg="儲存檔案失敗！", error=e)
                return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案
            else:
                logging_process_step(f"儲存檔案至路徑：{file_path}")
        except Exception as e:
            logging_exc_error(msg="儲存檔案失敗！", error=e)
            return EXIT_CODE_SAVE_FAILURE    # 作業異當終止：無法儲存檔案

        # if wb:
        #     xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留

    # =========================================================================
    # 結束程式
    # =========================================================================
    logging_process_step(f"《========== 程式終止執行：{program_name} ==========》")
    return EXIT_CODE_SUCCESS    # 作業正常結束


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)