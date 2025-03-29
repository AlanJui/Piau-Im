# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sqlite3
import sys
import unicodedata
from pathlib import Path

import xlwings as xw
from dotenv import load_dotenv

from a000_重置漢字標音工作表 import main as a000_main
from mod_excel_access import delete_sheet_by_name, ensure_sheet_exists
from mod_file_access import save_as_new_file
from mod_字庫 import JiKhooDict  # 漢字字庫物件
from mod_帶調符音標 import (
    apply_tone,
    cing_tu_khong_ze_ji_guan,
    clean_im_piau,
    fix_im_piau_spacing,
    handle_o_dot,
    is_han_ji,
    is_im_piau,
    read_text_with_han_ji,
    read_text_with_im_piau,
    separate_tone,
    tng_im_piau,
    tng_tiau_ho,
    tng_un_bu,
    zing_li_zuan_ku,
)
from mod_標音 import PiauIm  # 漢字標音物件
from mod_標音 import han_ji_piau_im  # 依據【台語音標】查找【漢字標音】
from mod_標音 import split_tai_gi_im_piau  # 分解台語音標
from mod_河洛話 import han_ji_ca_piau_im

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_NO_FILE = 90 # 無法找到檔案
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
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
)

init_logging()


# =========================================================================
# 程式區域函式
# =========================================================================
def read_han_ji_from_text_file(wb, filename:str, sheet_name:str='漢字注音', start_row:int=5):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    #------------------------------------------------------------------------------
    # 填入【漢字】
    #------------------------------------------------------------------------------
    text_with_han_ji = read_text_with_han_ji(filename=filename)

    row_han_ji = start_row      # 漢字位置
    start_col = 4   # 從D欄開始
    max_col = 18    # 最大可填入的欄位（R欄）

    col = start_col

    text = ""
    for han_ji_ku in text_with_han_ji:
        for han_ji in han_ji_ku:
            if col > max_col:
                # 超過欄位，換到下一組行
                row_han_ji += 4
                col = start_col

            text += han_ji
            sheet.cells(row_han_ji, col).value = han_ji
            sheet.cells(row_han_ji, col).select()  # 選取，畫面滾動
            col += 1  # 填入後右移一欄
            # 以下程式碼有假設：每組漢字之結尾，必有標點符號

        # 段落終結處：換下一段落
        if col > max_col:
            # 超過欄位，換到下一組行
            row_han_ji += 4
            col = start_col
        sheet.cells(row_han_ji, col).value = "=CHAR(10)"
        text += "\n"
        row_han_ji += 4
        col = start_col

    # 填入文章終止符號：φ
    sheet["V3"].value = text
    sheet.cells(row_han_ji, col).value = "φ"
    print(f"已將文章之漢字純文字檔讀入，並填進【{sheet_name}】工作表！")

    return text_with_han_ji


def cue_han_ji_piau_im(wb, text_with_han_ji:list) -> list:
    """查漢字讀音：依【漢字】查找【台語音標】，並依指定之【標音方法】輸出【漢字標音】"""
    #------------------------------------------------------------------------------
    # 籌備作業
    #------------------------------------------------------------------------------

    # 建置 PiauIm 物件，供作漢字拼音轉換作業
    han_ji_khoo_name = wb.names['漢字庫'].refers_to_range.value
    ue_im_lui_piat = wb.names['語音類型'].refers_to_range.value    # 指定【台語音標】轉換成【漢字標音】的方法
    piau_im_huat = wb.names['標音方法'].refers_to_range.value    # 指定【台語音標】轉換成【漢字標音】的方法
    piau_im = PiauIm(han_ji_khoo=han_ji_khoo_name)            # 指定漢字自動查找使用的【漢字庫】

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

    try:
        # 連接指定資料庫
        conn = sqlite3.connect(DB_HO_LOK_UE)
        cursor = conn.cursor()

        # 指定【漢字注音】工作表為【作用工作表】
        sheet = wb.sheets['漢字注音']
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

        im_piau_list = []
        idx = 0
        row = start_row
        col = start_col
        for han_ji_ku in text_with_han_ji:
            for han_ji in han_ji_ku:
                im_piau = ""
                if han_ji == 'φ':
                    EOF = True
                    msg = "【文字終結】"
                elif han_ji == '\n':
                    msg = "【換行】"
                elif han_ji == None or han_ji.strip() == "":  # 若儲存格內無值
                    msg = "【空缺】"    # 表【儲存格】未填入任何字/符，不同於【空白】字元
                else:
                    # 若不為【標點符號】，則以【漢字】處理
                    if not is_han_ji(han_ji):
                        im_piau = han_ji
                        msg = f"{han_ji}：略過不轉換！"
                    else:
                        # 自【漢字庫】查找作業
                        result = han_ji_ca_piau_im(
                            cursor=cursor,
                            han_ji=han_ji,
                            ue_im_lui_piat=ue_im_lui_piat)

                        # 若【漢字庫】查無此字，登錄至【缺字表】
                        if not result:
                            khuat_ji_piau_ji_khoo.add_or_update_entry(
                                han_ji=han_ji,
                                tai_gi_im_piau='N/A',
                                kenn_ziann_im_piau='N/A',
                                coordinates=(row, col)
                            )
                            msg = f"{han_ji}：查無此字！"
                        else:
                            # 依【漢字庫】查找結果，輸出【台語音標】和【漢字標音】
                            siann_bu = result[0]['聲母']
                            un_bu = result[0]['韻母']
                            un_bu = tng_un_bu(un_bu)
                            tiau_ho = result[0]['聲調']
                            if tiau_ho == "6":
                                # 若【聲調】為【6】，則將【聲調】改為【7】
                                tiau_ho = "7"
                            # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
                            im_piau = ''.join([siann_bu, un_bu, tiau_ho])

                            # 【標音字庫】添加或更新【漢字】資料
                            piau_im_ji_khoo.add_or_update_entry(
                                han_ji=han_ji,
                                tai_gi_im_piau=im_piau,
                                kenn_ziann_im_piau='N/A',
                                coordinates=(row, col)
                            )
                            msg = f"{han_ji}： [{im_piau}]"

                    # 顯示處理進度
                    im_piau_list.append(im_piau)
                    print(f"{idx}. ({row}, {col}) = {msg}")

                # 每處理一個字，右移一格
                idx += 1
                col += 1
                if col == end_col:
                    # 若已處理完一行，則換行
                    row += 4
                    col = start_col

            # 讀完一個段落
            im_piau_list.append("\n")
            print(f"({idx}. ({row}, {col}) = 【段落終結】")
            # 若已處理完一段落，則換行
            row += 4
            col = start_col
            idx += 1

        #----------------------------------------------------------------------
        # 作業結束前處理
        #----------------------------------------------------------------------
        # 關閉資料庫連線
        conn.close()
        logging_process_step("已完成【漢字】查找標音作業。")
        return im_piau_list
    except Exception as e:
        # 你可以在這裡加上紀錄或處理，例如:
        logging.exception("為【漢字】查找標音作業，發生異常狀況！")
        # 再次拋出異常，讓外層函式能捕捉
        raise
    finally:
        # 將【標音字庫】、【缺字表】字典，寫入 Excel 工作表
        khuat_ji_piau_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=khuat_ji_piau_name)
        piau_im_ji_khoo.write_to_excel_sheet(wb=wb, sheet_name=piau_im_sheet_name)
        # 關閉資料庫連線
        conn.close()


def fill_in_ping_im(wb, han_ji_list:list, im_piau_list:list, use_tiau_ho:bool=True, sheet_name:str='漢字注音', start_row:int=5, piau_im_soo_zai:int=-2):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    #------------------------------------------------------------------------------
    # 填入【音標】
    #------------------------------------------------------------------------------
    row_han_ji = start_row      # 漢字位置
    row_im_piau = row_han_ji + piau_im_soo_zai   # 標音所在: -1 ==> 自動標音； -2 ==> 人工標音
    start_col = 4   # 從D欄開始
    max_col = 18    # 最大可填入的欄位（R欄）

    #------------------------------------------------------------------------------
    # 填入【音標】
    #------------------------------------------------------------------------------
    col = start_col
    idx = 0
    # 執行到此，【音標】應已轉換為【帶調號之TLPA音標】
    for idx, han_ji in enumerate(han_ji_list):
        if col > max_col:   # 若已填滿一行（col = 19），則需換行
            row_han_ji += 4
            row_im_piau += 4
            col = start_col

        im_piau = ""
        msg = ""
        if han_ji == "\n":
            msg = "《換行》"
        else:
            tlpa_im_piau = im_piau_list[idx]
            if tlpa_im_piau == "":
                # 若音標為空白，表示遇有漢字未查找到音標
                msg = "《空白字元》"    # 標示為：【沒有音標】
            elif han_ji and is_han_ji(han_ji):
                # 若 cell_char 為漢字，
                im_piau = tlpa_im_piau
                msg = f"{han_ji} [ {im_piau} ] <-- {im_piau_list[idx]}"
            else:
                msg = f"{han_ji}《標點符號》"

        # 填入【音標】
        sheet.cells(row_im_piau, col).select()
        sheet.cells(row_im_piau, col).value = im_piau
        print(f"{idx}. ({row_im_piau}, {col})：{msg}")
        idx += 1
        col += 1
        if han_ji == "\n":
            # 若遇到換行符號，表示段落結束，換到下一段落
            row_han_ji += 4     # 漢字位置
            row_im_piau += 4    # 音標位置
            col = start_col     # 每句開始的欄位

    # 更新下一組漢字及TLPA標音之位置
    row_han_ji += 4     # 漢字位置
    row_im_piau += 4    # 音標位置
    col = start_col     # 每句開始的欄位

    print(f"已將【漢字】及【漢字標音】填入【{sheet_name}】工作表！")


def fill_in_jin_kang_ping_im(wb, han_ji_list:list, ping_im_file:str, use_tiau_ho:bool=True, sheet_name:str='漢字注音', start_row:int=5, piau_im_soo_zai:int=-2):
    # 讀取整篇文章之【標音文檔】
    text_with_im_piau = read_text_with_im_piau(filename=ping_im_file)

    #------------------------------------------------------------------------------
    # 填入【音標】
    #------------------------------------------------------------------------------
    fixed_im_piau_ku = []
    for im_piau_ku in text_with_im_piau:
        #------------------------------------------------------------------------------
        # 處理【音標】句子，將【字串】(String)資料轉換成【清單】(List)，使之與【漢字】一一對映
        #------------------------------------------------------------------------------
        # 整理整個句子，移除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
        im_piau_ku_cleaned = zing_li_zuan_ku(im_piau_ku)

        # 解構【音標】組成之【句子】，變成單一【帶調符音標】清單
        raw_im_piau_list = im_piau_ku_cleaned.split()

        # 查檢【音標】是否有【漢字+音標】之異常組合，若有則進行處理
        fixed_im_piau_list = []
        for i, raw_im_piau in enumerate(raw_im_piau_list):
            fixed_im_piau = fix_im_piau_spacing(raw_im_piau)
            # 可能因【校正音標】産生兩個音標，需分開處理
            fixed_im_piau_list.extend(fixed_im_piau.split())
            # print(f"已整理音標：{i}. {raw_im_piau} --> {fixed_im_piau}")

        # 轉換成【帶調號拼音】
        im_piau_list = []
        for i, im_piau in enumerate(fixed_im_piau_list):
            # 排除標點符號不進行韻母轉換
            if is_im_piau(im_piau):
                # 若為標點符號，無需轉換
                tlpa_im_piau = im_piau
            elif im_piau and is_han_ji(im_piau[0]):
                # 若為漢字，表示遇有漢字未查找到音標
                tlpa_im_piau = ""
            else:
                # 符合【帶調符音標】格式者，則進行【帶調號音標】轉換
                if im_piau == "":
                    tlpa_im_piau = ""
                else:
                    tlpa_im_piau = tng_im_piau(im_piau)    # 完成轉換之音標 = 音標帶調號
            im_piau_list.append(tlpa_im_piau)
            # print(f"已轉換音標：{i}. {im_piau} --> {tlpa_im_piau}")

        fixed_im_piau_ku.extend(im_piau_list)

    #------------------------------------------------------------------------------
    # 填入【音標】
    #------------------------------------------------------------------------------
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    row_han_ji = start_row      # 漢字位置
    start_col = 4   # 從D欄開始
    max_col = 18    # 最大可填入的欄位（R欄）

    col = start_col

    row_han_ji = start_row  # 重設【行數】為： 5（第5行）
    row_im_piau = row_han_ji + piau_im_soo_zai   # 標音所在: -1 ==> 自動標音； -2 ==> 人工標音
    im_piau_list = fixed_im_piau_ku
    col = start_col     # 重設【欄數】為： 4（D欄）
    im_piau_idx = 0
    # 執行到此，【音標】應已轉換為【帶調號之TLPA音標】
    while im_piau_idx < len(im_piau_list):
        if col > max_col:   # 若已填滿一行（col = 19），則需換行
            row_han_ji += 4
            row_im_piau += 4
            col = start_col
        han_ji = sheet.cells(row_han_ji, col).value
        if han_ji == "\n":
            # 若遇到換行符號，表示段落結束，換到下一段落
            row_han_ji += 4     # 漢字位置
            row_im_piau += 4    # 音標位置
            col = start_col     # 每句開始的欄位
            continue

        tlpa_im_piau = im_piau_list[im_piau_idx]
        im_piau = ""
        if tlpa_im_piau == "":
            # 若音標為空白，表示遇有漢字未查找到音標
            im_piau = ""    # 標示為：【沒有音標】
        elif han_ji and is_han_ji(han_ji):
            # 若 cell_char 為漢字，
            if use_tiau_ho:
                # 若設定【音標帶調號】，將 tlpa_word（音標），轉換音標格式為：【聲母】+【韻母】+【調號】
                im_piau = tng_tiau_ho(tlpa_im_piau)
            else:
                im_piau = tlpa_im_piau
        # 填入【音標】
        sheet.cells(row_im_piau, col).select()
        sheet.cells(row_im_piau, col).value = im_piau
        print(f"（{row_im_piau}, {col}）已填入: {han_ji} [ {im_piau} ] <-- {im_piau_list[im_piau_idx]}")
        im_piau_idx += 1
        col += 1

    # 更新下一組漢字及TLPA標音之位置
    row_han_ji += 4     # 漢字位置
    row_im_piau += 4    # 音標位置
    col = start_col     # 每句開始的欄位

    print(f"已將漢字及TLPA注音填入【{sheet_name}】工作表！")


# =========================================================================
# 主作業程序
# =========================================================================
def main():
    # 預設檔案名稱
    default_han_ji_file = "tmp_p1_han_ji.txt"
    # default_ping_im_file = "tmp_p2_ping_im.txt"
    default_ping_im_file = ""

    # 檢查是否有指定檔案名稱，若無則使用預設檔名
    # 命令列參數處理：sys.argv[1] = 漢字檔案, sys.argv[2] = 拼音檔案
    han_ji_file = sys.argv[1] if len(sys.argv) > 1 else default_han_ji_file
    ping_im_file = sys.argv[2] if len(sys.argv) > 2 else default_ping_im_file

    # 檢查是否有 'ho' 參數，若有則使用標音格式二：【聲母】+【韻母】+【調號】
    if "hu" in sys.argv:  # 若命令行參數包含 'bp'，則使用 BP
        use_tiau_ho = False
    else:
        use_tiau_ho = True
    # 以作用中的Excel活頁簿為作業標的
    wb = xw.apps.active.books.active
    if wb is None:
        logging.error("無法找到作用中的Excel活頁簿。")
        return

    # 備妥工作檔
    a000_main(wb)

    # 讀取整篇文章之【漢字】純文字檔案。
    text_with_han_ji = read_han_ji_from_text_file(wb, filename=han_ji_file)

    # 將 text_with_han_ji 整編為【漢字】列表
    han_ji_list = []
    for han_ji_ku in text_with_han_ji:
        for han_ji in han_ji_ku:
            han_ji_list.append(han_ji)
        # 段落終結處：換下一段落
        han_ji_list.append("\n")

    # 查找【漢字】之【音標】
    im_piau_list = cue_han_ji_piau_im(wb, text_with_han_ji)

    # 將【漢字】及【漢字標音】填入【漢字注音】工作表
    fill_in_ping_im(wb,
        han_ji_list=han_ji_list,
        im_piau_list=im_piau_list,
        use_tiau_ho=use_tiau_ho,
        start_row=5,
        piau_im_soo_zai=-1) # -1: 自動標音；-2: 人工標音

    # 將【漢字】及【標音文檔】填入【漢字注音】工作表
    if ping_im_file:
        fill_in_jin_kang_ping_im(wb,
            han_ji_list=han_ji_list,
            ping_im_file=ping_im_file,
            use_tiau_ho=use_tiau_ho,
            start_row=5,
            piau_im_soo_zai=-2) # -1: 自動標音；-2: 人工標音

    # 依【漢字】之【台語音標】，轉換成【漢字標音】
    han_ji_piau_im(wb)

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

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
