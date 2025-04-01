# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import re
import sys
import unicodedata
from pathlib import Path

import xlwings as xw
from dotenv import load_dotenv

from mod_file_access import save_as_new_file
from mod_帶調符音標 import (
    apply_tone,
    cing_tu_khong_ze_ji_guan,
    clean_im_piau,
    handle_o_dot,
    is_han_ji,
    is_im_piau,
    separate_tone,
    tng_im_piau,
    tng_tiau_ho,
    tng_un_bu,
    zing_li_zuan_ku,
)

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
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()


# =========================================================================
# 程式區域函式
# =========================================================================
def fix_im_piau_spacing(im_piau: str) -> str:
    """
    若音標的首字為漢字，則在首個漢字之後插入空白字元。
    如：「僇lâng」→「僇 lâng」
    """
    if im_piau and is_han_ji(im_piau[0]):
        return im_piau[0] + ' ' + im_piau[1:]
    return im_piau

# 用途：從純文字檔案讀取資料並回傳 [(漢字, TLPA), ...] 之格式
def read_text_with_han_ji(filename: str = "p2_han_ji.txt") -> list:
    text_with_han_ji = []
    with open(filename, "r", encoding="utf-8") as f:
        # 先移除 `\u200b`，確保不會影響 TLPA 拼音對應
        lines = [re.sub(r"[\u200b]", "", line.strip()) for line in f if line.strip()]

    for i in range(0, len(lines), 1):
        han_ji = lines[i]
        text_with_han_ji.append(han_ji)

    return text_with_han_ji


# 用途：從純文字檔案讀取資料並回傳 [(漢字, TLPA), ...] 之格式
def read_text_with_im_piau(filename: str = "ping_im.txt") -> list:
    text_with_tlpa = []
    with open(filename, "r", encoding="utf-8") as f:
        # 先移除 `\u200b`，確保不會影響 TLPA 拼音對應
        lines = [re.sub(r"[\u200b]", "", line.strip()) for line in f if line.strip() and not line.startswith("zh.wikipedia.org")]

    # for i in range(0, len(lines), 2):
    for i in range(0, len(lines), 1):
        im_piau = lines[i].replace("-", " ")  # 替換 "-" 為空白字元
        text_with_tlpa.append((im_piau))

    return text_with_tlpa

# =========================================================================
# 用途：將漢字及TLPA標音填入Excel指定工作表
# =========================================================================
def fill_han_ji_and_ping_im(wb, han_ji_filename:str, ping_im_filename:str, use_tiau_ho:bool=True, sheet_name:str='漢字注音', start_row:int=5, piau_im_soo_zai:int=-2):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    #------------------------------------------------------------------------------
    # 填入【漢字】
    #------------------------------------------------------------------------------
    text_with_han_ji = read_text_with_han_ji(filename=han_ji_filename)

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
        sheet.cells(row_han_ji, col).value = "=CHAR(10)"
        text += "\n"
        row_han_ji += 4
        col = start_col

    # 填入文章終止符號：φ
    sheet["V3"].value = text
    sheet.cells(row_han_ji, col).value = "φ"


    #------------------------------------------------------------------------------
    # 填入【音標】
    #------------------------------------------------------------------------------
    text_with_im_piau = read_text_with_im_piau(filename=ping_im_filename)

    row_han_ji = start_row      # 漢字位置
    start_col = 4   # 從D欄開始
    max_col = 18    # 最大可填入的欄位（R欄）

    col = start_col

    row_han_ji = start_row  # 重設【行數】為： 5（第5行）
    row_im_piau = row_han_ji + piau_im_soo_zai   # 標音所在: -1 ==> 自動標音； -2 ==> 人工標音
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
    default_ping_im_file = "tmp_p2_ping_im.txt"

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

    fill_han_ji_and_ping_im(wb,
        han_ji_filename=han_ji_file,
        ping_im_filename=ping_im_file,
        use_tiau_ho=use_tiau_ho,
        start_row=5,
        piau_im_soo_zai=-2) # -1: 自動標音；-2: 人工標音

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
