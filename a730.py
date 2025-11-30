import argparse
import re
import unicodedata
from typing import Optional, Tuple

import xlwings as xw

from mod_BP_tng_huan import (
    convert_bp_im_piau_to_zu_im,
    convert_siann_bu,
    convert_to_tiau_hu,
    convert_un_bu,
)
from mod_BP_tng_huan_ping_im import convert_TLPA_to_BP
from mod_TLPA_tng_BP import (
    convert_tlpa_to_zu_im_by_siann_bu,
    convert_tlpa_to_zu_im_by_tiau,
    convert_tlpa_to_zu_im_by_un_bu,
    convert_tlpa_to_zu_im_by_un_kap_tiau,
    kam_u_tiau_ho,
)

# from a720_製作注音打字練習工作表 import calculate_total_rows
from mod_帶調符音標 import kam_si_u_tiau_hu, tng_im_piau, tng_tiau_ho

# from dotenv import load_dotenv


# =========================================================================
# 常數定義
# =========================================================================
# 【漢字注音】工作表
START_COL = 'D'
END_COL = 'R'
BASE_ROW = 3
ROWS_PER_GROUP = 4

# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_FAILURE = 1  # 失敗
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_NO_FILE = 90 # 無法找到檔案
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()


# =========================================================================
# Excel 相關輔助函數
# =========================================================================
def calculate_total_rows(sheet, start_col=START_COL, end_col=END_COL, base_row=BASE_ROW, rows_per_group=ROWS_PER_GROUP):
    """Compute how many row groups exist based on the described worksheet layout."""
    total_rows = 0
    current_base = base_row

    while True:
        han_row = current_base + 2
        pronunciation_row = current_base + 3
        target_range = sheet.range(f'{start_col}{han_row}:{end_col}{pronunciation_row}')
        values = target_range.value

        if not _has_meaningful_data(values):
            break

        total_rows += 1
        current_base += rows_per_group

    return total_rows


def _has_meaningful_data(values):
    """Return True if any cell in the provided values contains non-blank data."""
    def _is_blank(cell):
        if cell is None:
            return True
        if isinstance(cell, str) and cell.strip() == '':
            return True
        return False

    if values is None:
        return False

    if not isinstance(values, list):
        return not _is_blank(values)

    for row in values:
        cells = row if isinstance(row, list) else [row]
        for cell in cells:
            if not _is_blank(cell):
                return True
    return False


def is_punctuation(char):
    """
    判斷是否為標點符號
    """
    if char is None or str(char).strip() == '':
        return False

    # 常見的中文標點符號
    chinese_punctuation = '，。！？；：「」『』（）【】《》〈〉、—…～'
    # 常見的英文標點符號
    english_punctuation = ',.!?;:"()[]{}/<>-_=+*&^%$#@`~|\\\'\"'

    return str(char) in chinese_punctuation or str(char) in english_punctuation


def is_line_break(char):
    """
    判斷是否為換行控制字元
    """
    if char is None:
        return False

    return char == '\n' or str(char).strip() == '' or char == 10

# def tlpa_tng_han_ji_piau_im(piau_im, piau_im_huat, tai_gi_im_piau):
#     tai_gi_im_piau_iong_tiau_ho = convert_tl_with_tiau_hu_to_tlpa(tai_gi_im_piau)
#     siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_im_piau_iong_tiau_ho)

#     if siann_bu == "" or siann_bu == None:
#         # siann_bu = "Ø"
#         siann_bu = 'ø'

#     ok = False
#     han_ji_piau_im = ""
#     try:
#         han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(
#             piau_im_huat=piau_im_huat,
#             siann_bu=siann_bu,
#             un_bu=un_bu,
#             tiau_ho=tiau_ho,
#         )
#         if han_ji_piau_im: # 傳回非空字串，表示【漢字標音】之轉換成功
#             ok = True
#         else:
#             logging_warning(f"【台語音標】：[{tai_gi_im_piau}]，轉換成【{piau_im_huat}漢字標音】拚音/注音系統失敗！")
#     except Exception as e:
#         logging_exception(f"piau_im.han_ji_piau_im_tng_huan() 發生執行時期錯誤: 【台語音標】：{tai_gi_im_piau}", e)
#         han_ji_piau_im = ""

#     # 若 ok 為 False，表示轉換失敗，則將【台語音標】直接傳回
#     return han_ji_piau_im


def decompose_bp_zu_im(bp_zu_im, tone_map_type='tlpa'):
    """
    將注音符號或羅馬拼音分解成個別字元

    Args:
        bp_zu_im (str): 注音符號

    Returns:
        list: 分解後的字元列表
    """

    # 【調符】與【調號】對映：hu_tiau_map
    # 【調號】與【按鍵】對映：tiau_key_map
    hu_tiau_map = {
        "˫": "陽去",   # 陽去
        "˪": "陰去",   # 陰去
        "ˋ": "上聲",  # 上聲
        "ˊ": "陽平",  # 陽平
        "˙": "陽入",   # 陽入
        "⁰": "輕聲",   # 輕聲
    }

    tiau_kian_map = {
        "陰平": ":",    # 陰平（無調號）
        "陽去": "5",   # 陽去
        "陰去": "3",   # 陰去
        "上聲": "4",  # 上聲
        "陽平": "6",  # 陽平
        "陰入": "[",    # 陰入（無調號）
        "陽入": "]",   # 陽入
        "輕聲": "7",   # 輕聲
    }

    # 方音符號處理
    chars = list(bp_zu_im)
    result = []
    okay = False
    counter = len(chars)
    index = 0

    for i, char in enumerate(chars):
        # hu_tiau_map = ['˪', '˫', 'ˋ', 'ˊ', '˙']
        if char in hu_tiau_map:
            # 是聲調符號，轉換為按鍵
            tiau_ho = hu_tiau_map[char]
            result.append(tiau_kian_map[tiau_ho])
            okay = True
        else:
            result.append(char)
        index += 1

    # 如果沒有聲調符號，則可能是：【陰平調】或【陰入調】
    if not okay:
        if index == counter:
            if chars[counter - 1] in ['ㆴ', 'ㆵ', 'ㆻ', 'ㆷ']:
                # 若最後一個字元，是【入聲韻尾】，則視為【陰入調】
                result.append(tiau_kian_map["陰入"])
            else:
                # 若最後一個字元，亦不是【入聲韻尾】，則視為【陰平調】
                result.append(tiau_kian_map["陰平"])

    return result



def process(tone_map_type: str) -> bool:
    """
    主處理函數
    :param tone_map_type: str - 聲調對照類型
    :return: bool - 處理是否成功
    """
    success = False

    # 取得目前作用中的 Excel 活頁簿
    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_exception("無法取得作用中的 Excel 活頁簿，請確認 Excel 是否已開啟且有作用中的活頁簿。", e)
        return False

    # 取得【漢字注音】工作表
    try:
        han_ji_zu_im_sheet = wb.sheets['漢字注音']
    except Exception as e:
        logging_exception("無法取得【漢字注音】工作表，請確認該工作表是否存在於活頁簿中。", e)
        return False

    # 取得或建立【打字練習表】工作表
    try:
        typing_sheet = wb.sheets['打字練習表']
        print("已找到【打字練習表】工作表")
    except Exception:
        typing_sheet = wb.sheets.add('打字練習表')
        print("已建立新的【打字練習表】工作表")

    # 清空打字練習表的內容（從第4行開始）
    # typing_sheet.range('B4:M2000').clear()
    typing_sheet.range('B4:M2000').clear_contents()

    #============================================================================
    # 開始處理資料
    #============================================================================

    # 開始處理資料
    current_row = 4  # 從第4行開始填入資料

    print("開始處理漢字注音資料...")

    # 根據【漢字注音】工作表，計算【總列數】
    # 第1列：{D3:R6} - 第3格D5, 第4格D6
    # 第2列：{D7:R10} - 第3格D9, 第4格D10
    # 第3列：{D11:R14} - 第3格D13, 第4格D14
    # 第4列：{D15:R18} - 第3格D17, 第4格D18
    # 第5列：{D19:R22} - 第3格D21, 第4格D22
    # ... 以此類推
    total_rows = calculate_total_rows(han_ji_zu_im_sheet)
    if total_rows == 0:
        print("【漢字注音】工作表沒有可用資料，結束處理")
        return success

    #----------------------------------------------------------------------------
    # 處理每一列資料
    #----------------------------------------------------------------------------
    print(f"總共需要處理 {total_rows} 列資料")

    # 計算各列的起始行號：3, 7, 11, 15, 19, 23
    row_starts = [3 + i * 4 for i in range(total_rows)]  # [3, 7, 11, 15, 19, 23]

    for row_group_index, base_row in enumerate(row_starts):
        # print(f"\n處理第 {row_group_index + 1} 列群組，基準行: {base_row}")
        print(f"\n----------------------------------------------------------")
        print(f"第 {row_group_index + 1} 列（漢字行: {base_row+2}）")
        print(f"----------------------------------------------------------")

        # 每列處理 D到R欄 (第4到第18欄)
        for col_index in range(4, 19):  # D(4) 到 R(18)
            try:
                col_letter = chr(64 + col_index)

                # 計算漢字和標音的實際行號
                han_ji_row = base_row + 2    # 第3格
                pronunciation_row = base_row + 3  # 第4格
                tai_gi_row = base_row + 1  # 第2格（目前未使用）

                # 取得當前單元格的資料
                han_ji = han_ji_zu_im_sheet.range(f'{col_letter}{han_ji_row}').value
                pronunciation = han_ji_zu_im_sheet.range(f'{col_letter}{pronunciation_row}').value
                tai_gi_piau_im = han_ji_zu_im_sheet.range(f'{col_letter}{tai_gi_row}').value

                # 檢查是否遇到終結符號
                if han_ji == 'φ':
                    print("    ==> 遇到終結符號，停止處理")
                    break

                # 檢查是否為換行控制字元
                if is_line_break(han_ji):
                    print(f"    ==> 欄位 {col_letter} 遇到換行控制字元，在打字練習表留空白行，跳至下一列")
                    # 留空白行（不填任何資料）
                    current_row += 1
                    # 跳出當前列的處理，進入下一列
                    break

                # 檢查是否為標點符號
                if is_punctuation(han_ji):
                    # print(f"    ==> 欄位 {col_letter} 是標點符號: {han_zi}")
                    # 標點符號只填入B欄，C欄及後續欄位留空
                    typing_sheet.range(f'B{current_row}').api.Value2 = str(han_ji)
                    current_row += 1
                    continue

                # 檢查資料是否有效
                if han_ji is None or pronunciation is None:
                    print(f"    ==> 欄位 {col_letter} 資料為空，跳過")
                    continue

                # 使用【閩拚音標】轉換成【注音符號】（以方音符號為基礎）當【漢字標音】
                if tone_map_type == 'bp' and tai_gi_piau_im is not None:
                    siann, un, tiau = convert_TLPA_to_BP(tai_gi_piau_im)

                    bp_im_piau = f"{siann}{un}{tiau}"
                    zu_im_siann, zu_im_un, zu_im_tiau_hu = convert_bp_im_piau_to_zu_im(bp_im_piau)
                    bp_zu_im = f"{zu_im_siann}{zu_im_un}{zu_im_tiau_hu}"
                    pronunciation = bp_zu_im

                # 使用【台語音標】轉換成【方音符號】當【漢字標音】
                if tone_map_type == 'tlpa' and tai_gi_piau_im is not None:
                    siann, un, tiau = split_tai_gi_im_piau(tai_gi_piau_im)

                    zu_im_siann = convert_tlpa_to_zu_im_by_siann_bu(siann)
                    zu_im_un = convert_tlpa_to_zu_im_by_un_kap_tiau(un, False)
                    # zu_im_un = convert_tlpa_to_zu_im_by_un_bu(un)
                    tiau_hu = convert_tlpa_to_zu_im_by_tiau(tiau)
                    zu_im = f"{zu_im_siann}{zu_im_un}{tiau_hu}"
                    pronunciation = zu_im

                # 填入純文字資料（不改變格式）
                typing_sheet.range(f'B{current_row}').api.Value2 = str(han_ji)
                typing_sheet.range(f'C{current_row}').api.Value2 = str(pronunciation)

                # 分解標音符號
                # tone_map_type = 'tfs'
                decomposed = decompose_bp_zu_im(str(pronunciation), tone_map_type)
                print(f"    ==> 鍵盤按鍵: {decomposed}\n")

                # 將分解後的字元填入 E~M 欄（純文字）
                for i, char in enumerate(decomposed):
                    if i < 9:  # 最多填入9個字元（E~M欄）
                        col_letter_target = chr(69 + i)  # E=69, F=70, ...
                        typing_sheet.range(f'{col_letter_target}{current_row}').api.Value2 = char

                # 顯示目前處理之【儲存格】位置與內容
                print(f"\n{col_index-3}.【{col_letter}{han_ji_row}】: 漢字={repr(han_ji)} [{tai_gi_piau_im}], 漢字標音={repr(pronunciation)}")
                current_row += 1
            except Exception as e:
                print(f"處理欄位 {col_letter} 時發生錯誤: {e}")
                continue

    success = True
    return success

def main():
    """
    主程式入口點
    """
    # 設定命令列參數解析
    success = False
    parser = argparse.ArgumentParser(description='自動製作打字練習表')
    parser.add_argument(
        'tone_map_type',
        nargs='?',
        default='tlpa',
        choices=['tlpa', 'bp'],
        help='聲調對照類型：roman (羅馬拼音，預設) 或 bp (閩拚)'
    )

    args = parser.parse_args()

    print("=== 自動製作打字練習表 ===")
    print(f"聲調對照類型: {args.tone_map_type}")
    print("請確保:")
    print("1. Excel 已開啟並有作用中的活頁簿")
    print("2. 活頁簿中包含【漢字注音】工作表")
    print("3. 漢字注音工作表的資料格式正確")
    print()

    success = process(args.tone_map_type)

    if success:
        print("\n✓ 處理作業成功！")
        return EXIT_CODE_SUCCESS
    else:
        print("\n✗ 處理作業失敗！")
        return EXIT_CODE_FAILURE


if __name__ == "__main__":
    main()