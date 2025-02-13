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
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import hong_im_tng_tai_gi_im_piau  # 方音符號轉台語音標
from mod_標音 import is_punctuation  # 是否為標點符號
from mod_標音 import siann_un_tiau_tng_piau_im  # 声、韻、調轉台語音標
from mod_標音 import split_hong_im_hu_ho  # 分解漢字標音
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_標音 import tai_gi_im_piau_tng_un_bu  # 台語音標轉韻部(方音轉強勢音)
from mod_標音 import tlpa_tng_han_ji_piau_im  # 台語音標轉台語音標

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
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def logging_process_step(msg):
    print(msg)
    logging.info(msg)

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_PROCESS_FAILURE = 3  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 作業程序
# =========================================================================
def ca_ji_kiat_ko_tng_piau_im(result, han_ji_khoo: str, piau_im: PiauIm, piau_im_huat: str):
    """查字結果出標音：查詢【漢字庫】取得之【查找結果】，將之切分：聲、韻、調"""
    if han_ji_khoo == "河洛話":
        #-----------------------------------------------------------------
        # 【白話音】：依《河洛話漢字庫》標注【台語音標】和【方音符號】
        #-----------------------------------------------------------------
        # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
        siann_bu = result[0]['聲母']
        un_bu = result[0]['韻母']
        un_bu = tai_gi_im_piau_tng_un_bu(un_bu)
        tiau_ho = result[0]['聲調']
        if tiau_ho == "6":
            # 若【聲調】為【6】，則將【聲調】改為【7】
            tiau_ho = "7"
    else:
        #-----------------------------------------------：------------------
        # 【文讀音】：依《廣韻字庫》標注【台語音標】和【方音符號】
        #-----------------------------------------------------------------
        siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(result[0]['標音'])
        if siann_bu == "" or siann_bu == None:
            siann_bu = "ø"

    # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
    # tai_gi_im_piau = siann_bu + un_bu + tiau_ho
    tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])

    # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
    if (piau_im_huat == "十五音" or piau_im_huat == "雅俗通") and (siann_bu == "" or siann_bu == None):
        siann_bu = "ø"
    han_ji_piau_im = siann_un_tiau_tng_piau_im(
        piau_im,
        piau_im_huat,
        siann_bu,
        un_bu,
        tiau_ho
    )
    return tai_gi_im_piau, han_ji_piau_im


def ca_han_ji_thak_im(wb, sheet_name='漢字注音', cell='V3', ue_im_lui_piat="白話音", han_ji_khoo="河洛話", db_name='Ho_Lok_Ue.db', module_name='mod_河洛話', function_name='han_ji_ca_piau_im'):
    """查漢字讀音：依【漢字】查找【台語音標】，並依指定之【標音方法】輸出【漢字標音】"""
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
    piau_im_sheet_name = '標音字庫'
    delete_sheet_by_name(wb=wb, sheet_name=piau_im_sheet_name)
    piau_im_ji_khoo = JiKhooDict()

    jin_kang_piau_im_sheet_name='人工標音字庫'
    delete_sheet_by_name(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)
    jin_kang_piau_im_ji_khoo = JiKhooDict()

    khuat_ji_piau_name = '缺字表'
    delete_sheet_by_name(wb=wb, sheet_name=khuat_ji_piau_name)
    khuat_ji_piau_ji_khoo = JiKhooDict()

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
            # 取得當前儲存格內含值
            han_ji_u_piau_im = False
            msg = ""
            cell = sheet.range((row, col))
            # 將文字顏色設為【自動】（黑色）
            cell.font.color = (0, 0, 0)  # 設定為黑色
            # 將儲存格的填滿色彩設為【無填滿】
            cell.color = None

            cell_value = cell.value
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
            else:
                # 若不為【標點符號】，則以【漢字】處理
                if is_punctuation(cell_value):
                    msg = f"{cell_value}"
                else:
                    # 查找漢字讀音
                    han_ji = cell_value
                    # 自【漢字庫】查找作業
                    result = han_ji_ca_piau_im(cursor=cursor,
                                                han_ji=han_ji,
                                                ue_im_lui_piat=ue_im_lui_piat)
                    # 若【漢字庫】查無此字，登錄至【缺字表】
                    if not result:
                        khuat_ji_piau_ji_khoo.add_or_update_entry(
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
                        # 【標音字庫】添加或更新【漢字】資料
                        piau_im_ji_khoo.add_or_update_entry(
                            han_ji=han_ji,
                            tai_gi_im_piau=tai_gi_im_piau,
                            kenn_ziann_im_piau='N/A',
                            coordinates=(row, col)
                        )
                        han_ji_u_piau_im = True

                    # 依據【人工標音】欄是否有輸入，決定【漢字標音】之處理方式
                    manual_input = sheet.range((row-2, col)).value
                    if manual_input:    # 若有人工輸入之處理作業
                        if '〔' in manual_input and '〕' in manual_input:
                            # 將人工輸入的〔台語音標〕轉換成【方音符號】
                            im_piau = manual_input.split('〔')[1].split('〕')[0]
                            tai_gi_im_piau = im_piau
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat,
                                tai_gi_im_piau=tai_gi_im_piau
                            )
                            han_ji_u_piau_im = True
                        elif '【' in manual_input and '】' in manual_input:
                            # 將人工輸入的【方音符號】轉換成【台語音標】
                            han_ji_piau_im = manual_input.split('【')[1].split('】')[0]
                            siann, un, tiau = split_hong_im_hu_ho(han_ji_piau_im)
                            # 依使用者指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            tai_gi_im_piau = hong_im_tng_tai_gi_im_piau(
                                siann=siann,
                                un=un,
                                tiau=tiau,
                                cursor=cursor,
                            )['台語音標']
                            han_ji_u_piau_im = True
                        else:
                            # 將人工輸入的【台語音標】，解構為【聲母】、【韻母】、【聲調】
                            tai_gi_im_piau = manual_input
                            # 依指定之【標音方法】，將【台語音標】轉換成其所需之【漢字標音】
                            han_ji_piau_im = tlpa_tng_han_ji_piau_im(
                                piau_im=piau_im,
                                piau_im_huat=piau_im_huat,
                                tai_gi_im_piau=tai_gi_im_piau
                            )
                            han_ji_u_piau_im = True

                        # 將人工輸入的【台語音標】置入【破音字庫】Dict
                        jin_kang_piau_im_ji_khoo.add_or_update_entry(
                            han_ji=han_ji,
                            tai_gi_im_piau=tai_gi_im_piau,
                            kenn_ziann_im_piau='N/A',
                            coordinates=(row, col)
                        )

                if han_ji_u_piau_im:
                    sheet.range((row - 1, col)).value = tai_gi_im_piau
                    sheet.range((row + 1, col)).value = han_ji_piau_im
                    if manual_input:
                        sheet.range((row, col)).font.color = (255, 0, 0)    # 將文字顏色設為【紅色】
                        sheet.range((row, col)).color = (255, 255, 0)       # 將底色設為【黄色】
                    msg = f"{han_ji}： [{tai_gi_im_piau}] /【{han_ji_piau_im}】"

            # 顯示處理進度
            col_name = xw.utils.col_name(col)   # 取得欄位名稱
            print(f"({row}, {col_name}) = {msg}")

            # 若讀到【換行】或【文字終結】，跳出逐欄取字迴圈
            if msg == "【換行】" or EOF:
                break

        # =================================================================
        # 每當處理一行 15 個漢字後，亦換到下一行
        if col == end_col - 1: print('\n')
        line += 1
        # 若已到【結尾】或【超過總行數】，則跳出迴圈
        if EOF or line > TOTAL_LINES:
            # 將【標音字庫】、【人工標音字庫】、【缺字表】三個字典，寫入 Excel 工作表
            piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
            jin_kang_piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=jin_kang_piau_im_sheet_name)
            khuat_ji_piau_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=khuat_ji_piau_name)
            break

    #----------------------------------------------------------------------
    # 作業結束前處理
    #----------------------------------------------------------------------
    # 關閉資料庫連線
    conn.close()
    print("已完成【台語音標】和【漢字標音】標注工作。")
    return EXIT_CODE_SUCCESS


def process(wb):
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

    if han_ji_khoo_name == "河洛話" and ue_im_lui_piat == "白話音":
        ca_han_ji_thak_im(
            wb=wb,
            sheet_name="漢字注音",
            cell="V3",
            ue_im_lui_piat=ue_im_lui_piat,  # "白話音"
            han_ji_khoo=han_ji_khoo_name,   # "河洛話",
            db_name=db_name,                # "Ho_Lok_Ue.db",
            module_name=module_name,        # "mod_河洛話",
            function_name=function_name     # "han_ji_ca_piau_im",
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
            function_name=function_name     # "han_ji_ca_piau_im",
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
        print(msg)
        logging.error(msg)
        return EXIT_CODE_INVALID_INPUT

    #--------------------------------------------------------------------------
    # 作業結尾處理
    #--------------------------------------------------------------------------
    # 要求畫面回到【漢字注音】工作表
    wb.sheets['漢字注音'].activate()
    # 儲存檔案
    save_as_new_file(wb=wb)
    logging.info("己存檔至路徑：{file_path}")
    return EXIT_CODE_SUCCESS


# =============================================================================
# 程式主流程
# =============================================================================
def main():
    # =========================================================================
    # 開始作業
    # =========================================================================
    logging.info("作業開始")

    # =========================================================================
    # (1) 取得專案根目錄。
    # =========================================================================
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
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
        result_code = process(wb)
        if result_code != EXIT_CODE_SUCCESS:
            logging_process_step("作業異常終止！")
            return result_code

    except Exception as e:
        print(f"作業過程發生未知的異常錯誤: {e}")
        logging.error(f"作業過程發生未知的異常錯誤: {e}", exc_info=True)
        return EXIT_CODE_UNKNOWN_ERROR

    finally:
        if wb:
            # xw.apps.active.quit()  # 確保 Excel 被釋放資源，避免開啟殘留
            logging.info("a702_查找及填入漢字標音.py 程式已執行完畢！")

    # =========================================================================
    # 結束作業
    # =========================================================================
    logging.info("作業完成！")
    return EXIT_CODE_SUCCESS


if __name__ == "__main__":
    exit_code = main()
    if exit_code == EXIT_CODE_SUCCESS:
        print("程式正常完成！")
    else:
        print(f"程式異常終止，錯誤代碼為: {exit_code}")
    sys.exit(exit_code)