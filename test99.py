import argparse
import re
import unicodedata
from typing import Optional, Tuple

import xlwings as xw

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


#====================================================================
# 韻母轉換函數
#====================================================================
def un_bu_tng_huan(un_bu: str) -> str:
    """
    將輸入的韻母依照轉換字典進行轉換
    :param un_bu: str - 韻母輸入
    :return: str - 轉換後的韻母結果
    """

    # 韻母起頭替換規則（優先進行）
    if un_bu.startswith("oa"):
        # 白話字：oa ==> 閩南語：ua
        un_bu = "u" + un_bu[1:]  # oan → uan, oann → uann
    elif un_bu.startswith("oe"):
        # 白話字：oe ==> 閩南語：ue
        un_bu = "u" + un_bu[1:]  # oe → ue, oeh → ueh

    # 韻母轉換字典
    un_bu_tng_huan_map_dict = {
        'ee': 'e',          # ee（ㄝ）= [ɛ]
        'er': 'e',          # er（ㄜ）= [ə]
        'erh': 'eh',        # er（ㄜ）= [ə]
        'or': 'o',          # or（ㄜ）= [ə]
        'ere': 'ue',        # ere = [əe]
        'ereh': 'ueh',      # ereh = [əeh]
        'ir': 'i',          # ir（ㆨ）= [ɯ] / [ɨ]
        'eng': 'ing',       # 白話字：eng ==> 閩南語：ing
        'oai': 'uai',       # 白話字：oai ==> 閩南語：uai
        'ei': 'e',          # 雅俗通十五音：稽
        'ou': 'oo',         # 雅俗通十五音：沽
        # 'onn': 'oonn',      # 雅俗通十五音：扛
        'uei': 'ue',        # 雅俗通十五音：檜
        'ueinn': 'uenn',    # 雅俗通十五音：檜
        'ur': 'u',          # 雅俗通十五音：艍
        'eng': 'ing',       # 白話字：eng ==> 台語音標：ing
        'ek': 'ik',         # 白話字：ek ==> 台語音標：ik
        'o͘': 'oo',          # 白話字：o͘ (o + U+0358) ==> 台語音標：oo
        'ⁿ': 'nn',          # 白話字：ⁿ ==> 台語音標：nn
    }

    # 韻母轉換，若不存在於字典中則返回原始韻母
    return un_bu_tng_huan_map_dict.get(un_bu, un_bu)

#====================================================================
# 【台語音標】韻母轉換函數
#====================================================================
def tai_gi_im_piau_tng_un_bu(tai_gi_im_piau: str) -> str:
    """
    將輸入的整體【台語音標】依韻母轉換字典進行韻母轉換
    :param tai_gi_im_piau: str - 整體台語音標 (例如 "kere1")
    :return: str - 韻母轉換後的台語音標 (例如 "kue1")
    """

    # 使用正則表達式拆解聲母、韻母、聲調
    match = re.match(r"([ptkghmnzcsjlrw]?)([a-z]+)(\d?)", tai_gi_im_piau, re.I)
    if match:
        siann_bu = match.group(1)  # 聲母
        un_bu = match.group(2)     # 韻母
        tiau_ho = match.group(3)   # 聲調

        # 韻母轉換字典
        un_bu_tng_huan_map_dict = {
            'ee': 'e',          # ee（ㄝ）= [ɛ]
            'er': 'e',          # er（ㄜ）= [ə]
            'erh': 'eh',        # er（ㄜ）= [ə]
            'or': 'o',          # or（ㄜ）= [ə]
            'ere': 'ue',        # ere = [əe]
            'ereh': 'ueh',      # ereh = [əeh]
            'ir': 'i',          # ir（ㆨ）= [ɯ] / [ɨ]
            'eng': 'ing',       # 白話字：eng ==> 閩南語：ing
            'oa': 'ua',         # 白話字：oa ==> 閩南語：ua
            'oe': 'ue',         # 白話字：oe ==> 閩南語：ue
            'oai': 'uai',       # 白話字：oai ==> 閩南語：uai
            'ei': 'e',          # 雅俗通十五音：稽
            'ou': 'oo',         # 雅俗通十五音：沽
            # 'onn': 'oonn',      # 雅俗通十五音：扛
            'uei': 'ue',        # 雅俗通十五音：檜
            'ueinn': 'uenn',    # 雅俗通十五音：檜
            'ur': 'u',          # 雅俗通十五音：艍
            'oa': 'ua',         # 白話字：oa ==> 台語音標：ua
            'oe': 'ue',         # 白話字：oe ==> 台語音標：ue
            'eng': 'ing',       # 白話字：eng ==> 台語音標：ing
            'ek': 'ik',         # 白話字：ek ==> 台語音標：ik
            'o͘': 'oo',          # 白話字：o͘ (o + U+0358) ==> 台語音標：oo
            'ⁿ': 'nn',          # 白話字：ⁿ ==> 台語音標：nn
        }

        # 韻母轉換
        converted_un_bu = un_bu_tng_huan_map_dict.get(un_bu, un_bu)

        # 合併轉換後的台語音標
        converted_tai_gi_im_piau = f"{siann_bu}{converted_un_bu}{tiau_ho}"
        return converted_tai_gi_im_piau

    # 若無法解析，返回原始輸入
    return tai_gi_im_piau

# ============================================================================
# 將使用【上標數字】表示的【調號】，轉換為普通數字
# ============================================================================
def replace_superscript_digits(input_str):
    # 上標數字與普通數字的映射字典
    superscript_digit_mapping = {
        '⁰': '0',
        '¹': '1',
        '²': '2',
        '³': '3',
        '⁴': '4',
        '⁵': '5',
        '⁶': '6',
        '⁷': '7',
        '⁸': '8',
        '⁹': '9',
    }
    return ''.join(superscript_digit_mapping.get(char, char) for char in input_str)

# ============================================================================
# 將【台語音標】分解為【聲母】、【韻母】、【調號】
# ============================================================================
def split_tai_gi_im_piau(im_piau: str, po_ci: bool = False):
    # 如果輸入之【音標】為【帶調符音標】，則需確保轉換為【帶調號TLPA音標】
    if kam_si_u_tiau_hu(im_piau):
        im_piau = tng_im_piau(im_piau)
        im_piau = tng_tiau_ho(im_piau)
    # 將輸入的台語音標轉換為小寫
    im_piau = im_piau.lower()
    # 查檢【台語音標】是否符合【標準】=【聲母】+【韻母】+【調號】
    tiau = im_piau[-1]
    tiau = replace_superscript_digits(str(tiau))

    # 矯正未標明陰平/陰入調號的情況
    if tiau in ['p', 't', 'k', 'h']:
        tiau = '4'
        im_piau += tiau
    elif tiau in ['a', 'e', 'i', 'o', 'u', 'm', 'n', 'g']:
        tiau = '1'
        im_piau += tiau

    # 聲母相容性轉換
    if im_piau.startswith("tsh"):
        im_piau = im_piau.replace("tsh", "c", 1)
    elif im_piau.startswith("ts"):
        im_piau = im_piau.replace("ts", "z", 1)

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")
    un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

    result = []

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_bu_as_m_or_ng_pattern.match(im_piau):
        siann_bu = ""
        un_bu = im_piau[:-1]
        tiau = im_piau[-1]
    else:
        siann_bu_match = siann_bu_pattern.match(im_piau)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()
            un_bu = im_piau[len(siann_bu):-1]
        else:
            siann_bu = ""
            un_bu = im_piau[:-1]

    # 轉換韻母
    un_bu = un_bu_tng_huan(un_bu)

    # 調整聲母大小寫
    if po_ci:
        siann_bu = siann_bu[0].upper() + siann_bu[1:] if siann_bu else ""

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

def split_hong_im_hu_ho(hong_im_piau_im):
    # 定義調符對應的字典
    Hong_Im_Tiau_Hu_Dict = {
        "ˋ": 2,
        "˪": 3,
        "ˊ": 5,
        "˫": 7,
        "\u02D9": 8,  # '˙'
    }

    # 編譯調符的正則表達式模式
    Hong_Im_Tiau_Hu = re.compile(r"[ˋ˪ˊ˫˙]", re.I)

    # 入声韻尾
    Jip_Siann_Un_Bue = {'ㆴ', 'ㆵ', 'ㆻ', 'ㆷ'}

    # 定義聲母的集合
    Siann_Mu_Ji = {
        'ㄅ', 'ㄆ', 'ㆠ', 'ㄇ',
        'ㄉ', 'ㄊ', 'ㄋ', 'ㄌ',
        'ㄍ', 'ㄎ', 'ㆣ', 'ㄏ', 'ㄫ',
        'ㄗ', 'ㄘ', 'ㆡ', 'ㄙ',
        'ㄐ', 'ㄑ', 'ㆢ', 'ㄒ',
        'ㄓ', 'ㄔ', 'ㄕ', 'ㄖ',
        'ㄭ', 'ㄪ', 'ㄬ', 'ㄈ',
    }

    # 步驟一：辨識【方音標音】中的【調號】值，並輸出僅有【聲母】和【韻母】的【無調號方音標音】
    if Hong_Im_Tiau_Hu.match(hong_im_piau_im[-1]):
        # 有【声調符號】的【方音標音】
        tiau_fu = hong_im_piau_im[-1]
        tiau_ho = Hong_Im_Tiau_Hu_Dict[tiau_fu]
        # 無【調號】，只有【聲母】和【韻母】的【方音標音】
        siann_ka_un_piau_im = hong_im_piau_im[:-1]
    else:
        # 無【声調符號】的【方音標音】，其【調號】值可為1或4，需依【方音標音】尾字，進行【辨識】
        if hong_im_piau_im[-1] in Jip_Siann_Un_Bue:
            # 【方音標音】最後一字為【入聲韻尾】，其【調號】值為4
            tiau_ho = 4
        else:
            # 【方音標音】最後一字非【入聲韻尾】，其【調號】值為1
            tiau_ho = 1
        siann_ka_un_piau_im = hong_im_piau_im

    # 步驟四：自【無調號方音標音】，取【聲母】和【韻母】
    if siann_ka_un_piau_im and siann_ka_un_piau_im[0] in Siann_Mu_Ji:
        # 有【聲母】
        siann_mu = siann_ka_un_piau_im[0]
        un_mu = siann_ka_un_piau_im[1:]
    else:
        # 無【聲母】
        siann_mu = ''
        un_mu = siann_ka_un_piau_im

    return [siann_mu, un_mu, str(tiau_ho)]


# =========================================================================
# 將首字母為大寫之羅馬拼音字母轉換為小寫（只處理第一個字母）
# =========================================================================
def normalize_im_piau_case(im_piau: str) -> str:
    im_piau = unicodedata.normalize("NFC", im_piau)  # 先標準化 Unicode
    return im_piau[0].lower() + im_piau[1:] if im_piau else im_piau

# =========================================================================
# 台語音標轉換為【漢字標音】之注音符號或羅馬字音標
# =========================================================================
def convert_tl_without_tiau_hu(tai_lo: str) -> str:
    """
    將帶有聲調符號的台羅拼音轉換為改良式【台語音標】（TLPA+）。
    """
    # **重要**：先將字串標準化為 NFC 格式，統一處理 Unicode 差異
    tai_lo = unicodedata.normalize("NFC", tai_lo)

    tone_number = ""

    # 1. 先處理聲調轉換
    for tone_mark, (base_char, number) in tone_mapping.items():
        if tone_mark in tai_lo:
            tai_lo = tai_lo.replace(tone_mark, base_char)  # 移除調號，還原原始母音
            tone_number = number  # 記錄對應的聲調數字
            break  # 只會有一個聲調符號，找到就停止

    # 2. 若有聲調數字，則加到末尾
    if tone_number:
        return tai_lo + tone_number

    return tai_lo  # 若無聲調符號則不變更

def convert_tl_to_tlpa(tai_lo: str) -> Optional[str]:
    """
    轉換台羅（TL）為台語音標（TLPA），只在單字邊界進行替換。
    """
    if not tai_lo:
        return None

    # 查檢【台語音標】是否符合【標準】=【聲母】+【韻母】+【調號】；若是將：【陰平】、【陰入】調，
    # 略去【調號】數值：1、4，則進行矯正
    # 先將傳入之【台語音標】的最後一個字元視作【調號】取出
    tiau = tai_lo[-1]
    # 若【調號】數值，使用上標數值格式，則替換為 ASCII 數字
    tiau = replace_superscript_digits(str(tiau))

    # 若輸入之【台語音標】未循【標準】，對【陰平】、【陰入】聲調，省略【調號】值：【1】/【4】
    # 則依此規則進行矯正：若【調號】（即：拼音最後一個字母）為 [ptkh]，則更正調號值為 4；
    # 則【調號】填入【韻母】之拼音字元，則將【調號】則更正為 1
    if tiau in ['p', 't', 'k', 'h']:
        tiau = '4'  # 聲調值為 4（陰入聲）
        tai_lo += tiau  # 為輸入之簡寫【台語音標】，添加【調號】
    elif tiau in ['a', 'e', 'i', 'o', 'u', 'm', 'n', 'g']:  # 如果最後一個字母是英文字母
        tiau = '1'  # 聲調值為 1（陰平聲）
        tai_lo += tiau  # 為輸入之簡寫【台語音標】，添加【調號】

    # 將【白話字】聲母轉換成【台語音標】（將 chh 轉換為 c；將 ch 轉換為 z）
    tai_lo = re.sub(r'^chh', 'c', tai_lo)  # `^` 表示「字串開頭」
    tai_lo = re.sub(r'^ch', 'z', tai_lo)  # `^` 表示「字串開頭」

    # 將【台羅音標】聲母轉換成【台語音標】（將 tsh 轉換為 c；將 ts 轉換為 z）
    tai_lo = re.sub(r'^tsh', 'c', tai_lo)  # `^` 表示「字串開頭」
    tai_lo = re.sub(r'^ts', 'z', tai_lo)  # `^` 表示「字串開頭」

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")

    # 韻母為 m 或 ng 這種情況的正規表示式 (m\d 或 ng\d)
    un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_bu_as_m_or_ng_pattern.match(tai_lo):
        siann_bu = ""  # 沒有聲母
        un_bu = tai_lo[:-1]  # 韻母是 m 或 ng
        tiau = tai_lo[-1]  # 聲調是最後一個字符
    else:
        # 使用正規表示式來匹配聲母
        siann_bu_match = siann_bu_pattern.match(tai_lo)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()  # 找到聲母
            un_bu = tai_lo[len(siann_bu):-1]  # 韻母部分
        else:
            siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
            un_bu = tai_lo[:-1]  # 韻母是剩下的部分，去掉最後的聲調

    # 轉換韻母
    un_bu = un_bu_tng_huan(un_bu)

    tai_gi = ''.join([siann_bu, un_bu, tiau])
    return tai_gi


def convert_tl_with_tiau_hu_to_tlpa(im_piau: str) -> Optional[str]:
    """
    將帶有聲調符號的台羅拼音轉換為改良式【台語音標】（TLPA+）。
    """
    # 1. 將首字母為大寫之羅馬拼音字母轉換為小寫（只處理第一個字母）
    im_piau = normalize_im_piau_case(im_piau)

    # 2. 先處理聲調轉換
    tai_lo_bo_taiu_hu = convert_tl_without_tiau_hu(im_piau)

    # 3. 將聲母轉換為 TLPA+
    tai_gi_im_piau = convert_tl_to_tlpa(tai_lo_bo_taiu_hu)

    return tai_gi_im_piau


def tlpa_tng_han_ji_piau_im(piau_im, piau_im_huat, tai_gi_im_piau):
    tai_gi_im_piau_iong_tiau_ho = convert_tl_with_tiau_hu_to_tlpa(tai_gi_im_piau)
    siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_im_piau_iong_tiau_ho)

    if siann_bu == "" or siann_bu == None:
        # siann_bu = "Ø"
        siann_bu = 'ø'

    ok = False
    han_ji_piau_im = ""
    try:
        han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(
            piau_im_huat=piau_im_huat,
            siann_bu=siann_bu,
            un_bu=un_bu,
            tiau_ho=tiau_ho,
        )
        if han_ji_piau_im: # 傳回非空字串，表示【漢字標音】之轉換成功
            ok = True
        else:
            logging_warning(f"【台語音標】：[{tai_gi_im_piau}]，轉換成【{piau_im_huat}漢字標音】拚音/注音系統失敗！")
    except Exception as e:
        logging_exception(f"piau_im.han_ji_piau_im_tng_huan() 發生執行時期錯誤: 【台語音標】：{tai_gi_im_piau}", e)
        han_ji_piau_im = ""

    # 若 ok 為 False，表示轉換失敗，則將【台語音標】直接傳回
    return han_ji_piau_im

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

                # 填入純文字資料（不改變格式）
                typing_sheet.range(f'B{current_row}').api.Value2 = str(han_ji)
                typing_sheet.range(f'C{current_row}').api.Value2 = str(pronunciation)

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