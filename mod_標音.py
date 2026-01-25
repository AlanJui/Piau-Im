import os
import re
import sqlite3
import unicodedata
from typing import Optional

from dotenv import load_dotenv

from mod_BP_tng_huan import (
    convert_bp_siann_un_tiau_to_zu_im,
)

# 將【台語音標】轉換成【閩拼音標】
from mod_BP_tng_huan_ping_im import (
    convert_TLPA_to_BP,
    convert_TLPA_to_BP_with_tone_marks,
)

# 將 TLPA+ 【台語音標】轉換成 MPS2 【台語注音二式】
from mod_convert_TLPA_to_MPS2 import convert_TLPA_to_MPS2
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_warning,
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

# 設定標點符號過濾
PUNCTUATIONS = (",", ".", "?", "!", ":", ";", "\u200B")

# 調符對映表（帶調符之元音/韻化輔音 → 不帶調符之拼音字母、調號數值）
tiau_hu_mapping = {
    "á": ("a", "2"), "à": ("a", "3"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "a̍": ("a", "8"), "a̋": ("a", "9"),
    "Á": ("A", "2"), "À": ("A", "3"), "Â": ("A", "5"), "Ǎ": ("A", "6"), "Ā": ("A", "7"), "A̍": ("A", "8"), "A̋": ("A", "9"),
    "é": ("e", "2"), "è": ("e", "3"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "e̍": ("e", "8"), "e̋": ("e", "9"),
    "É": ("E", "2"), "È": ("E", "3"), "Ê": ("E", "5"), "Ě": ("E", "6"), "Ē": ("E", "7"), "E̍": ("E", "8"), "E̋": ("E", "9"),
    "í": ("i", "2"), "ì": ("i", "3"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "i̍": ("i", "8"), "i̋": ("i", "9"),
    "Í": ("I", "2"), "Ì": ("I", "3"), "Î": ("I", "5"), "Ǐ": ("I", "6"), "Ī": ("I", "7"), "I̍": ("I", "8"), "I̋": ("I", "9"),
    "ó": ("o", "2"), "ò": ("o", "3"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "o̍": ("o", "8"), "ő": ("o", "9"),
    "Ó": ("O", "2"), "Ò": ("O", "3"), "Ô": ("O", "5"), "Ǒ": ("O", "6"), "Ō": ("O", "7"), "O̍": ("O", "8"), "Ő": ("O", "9"),
    "ú": ("u", "2"), "ù": ("u", "3"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "u̍": ("u", "8"), "ű": ("u", "9"),
    "Ú": ("U", "2"), "Ù": ("U", "3"), "Û": ("U", "5"), "Ǔ": ("U", "6"), "Ū": ("U", "7"), "U̍": ("U", "8"), "Ű": ("U", "9"),
    "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̍": ("m", "8"), "m̋": ("m", "9"),
    "Ḿ": ("M", "2"), "M̀": ("M", "3"), "M̂": ("M", "5"), "M̌": ("M", "6"), "M̄": ("M", "7"), "M̍": ("M", "8"), "M̋": ("M", "9"),
    "ń": ("n", "2"), "ǹ": ("n", "3"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̍": ("n", "8"), "n̋": ("n", "9"),
    "Ń": ("N", "2"), "Ǹ": ("N", "3"), "N̂": ("N", "5"), "Ň": ("N", "6"), "N̄": ("N", "7"), "N̍": ("N", "8"), "N̋": ("N", "9"),
}

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
    'ek': 'ik',         # 白話字：ek ==> 台語音標：ik
    'o͘': 'oo',          # 白話字：o͘ (o + U+0358) ==> 台語音標：oo
    'ⁿ': 'nn',          # 白話字：ⁿ ==> 台語音標：nn
}

# =========================================================================
# helper functions:  與 mod_帶調號母音轉換.py 重複，可考慮整合
# =========================================================================
def kam_si_u_tiau_hu(im_piau: str) -> bool:
    """是否有調符：判斷傳入之音標是否為【帶調符音標】

    Args:
        im_piau (str): 音標

    Returns:
        bool: [True] 帶調符音標；[False] 無調符音標
    """

    # 若【音標】末端為數值，表音標已是【帶調號音標】，直接回傳【無調符音標】
    # u_tiau_hu = False
    if im_piau[-1] in "123456789":
        return False

    # 將傳入【音標】字串，以標準化組合格式：NFC，將【帶調符拼音字母】標準化；
    # 令以下之處理作業，不會發生【看似相同】的【帶調符拼音字母】，其實使用
    # 不同之 Unicode 編碼
    im_piau = unicodedata.normalize("NFC", im_piau)

    #--------------------------------------------------------------------------------
    # 以【元音及韻化輔音清單】，比對傳入之【音標】，找出對應之【基本拼音字母】與【調號】
    #--------------------------------------------------------------------------------
    number = "1"  # 明確初始化 number 變數，以免未設定而發生錯誤
    for tone_mark, (base_char, number) in tiau_hu_mapping.items():
        if tone_mark in im_piau:
            # 轉換成【無調符音標】
            bo_tiau_hu_im_piau = im_piau.replace(tone_mark, base_char)  # noqa: F841
            break
    else:
        number = "1"  # 若沒有任何調符，number強制為1

    # 若 number 有值，且在 ["2", "3", "5", "6", "7", "8", "9"] 之中，則為【帶調符音標】
    if number in ["2", "3", "5", "6", "7", "8", "9"]:
        return True

    # 若【無調符音標】末端【拼音字母】為【hptk】之一，則為【陰入調】，則為【帶調符音標】
    if number == '1':
        if im_piau[-1] in "hptk":
            # 【無調符音標】末端為【hptk】之一，則為【陰入調】，聲調值為 4
            return True
        elif im_piau[-1] in "aeioumngAEIOUMN":
            # 【無調符音標】末端非【hptk】之一，則為【陰平調】，聲調值為 1
            return True

    return False

def tng_im_piau(im_piau: str, po_ci_tai_sia: bool = False, kan_hua: bool = False) -> str:
    """
    將【帶調符音標】（台羅拼音/台語音標）轉換成【帶調號TLPA音標】
    :param im_piau: str - 帶調符音標
    :param po_ci_tai_sia: bool - 是否保留【音標】之首字母大寫
    :param kan_hua: bool - 簡化：若是【簡化】，聲調值為 1 或 4 ，去除調號值
    :return: str - 轉換後的【帶調號TLPA音標】
    """
    # 遇標點符號，不做轉換處理，直接回傳
    if im_piau[-1] in PUNCTUATIONS:
        return im_piau

    #---------------------------------------------------------
    # 更換【音標】之【聲母】
    #---------------------------------------------------------
    # 將傳入【音標】字串，以標準化之 NFC 組合格式，調整【帶調符拼音字母】；
    # 令以下之處理作業，不會發生【看似相同】的【帶調符拼音字母】，其實使用
    # 不同之 Unicode 編碼
    im_piau = unicodedata.normalize("NFC", im_piau)

    if im_piau.startswith("tsh"):
        im_piau = im_piau.replace("tsh", "c", 1)
    elif im_piau.startswith("Tsh"):
        im_piau = im_piau.replace("Tsh", "C", 1)
    elif im_piau.startswith("ts"):
        im_piau = im_piau.replace("ts", "z", 1)
    elif im_piau.startswith("Ts"):
        im_piau = im_piau.replace("Ts", "Z", 1)
    elif im_piau.startswith("chh"):
        im_piau = im_piau.replace("chh", "c", 1)
    elif im_piau.startswith("Chh"):
        im_piau = im_piau.replace("Chh", "C", 1)
    elif im_piau.startswith("ch"):
        im_piau = im_piau.replace("ch", "z", 1)
    elif im_piau.startswith("Ch"):
        im_piau = im_piau.replace("Ch", "Z", 1)

    #---------------------------------------------------------
    # 更換【音標】之【韻母】
    #---------------------------------------------------------
    su_ji = im_piau[0]      # 保存【音標】之拼音首字母
    org_im_piau = im_piau
    im_piau = org_im_piau.lower()

    # 轉換【鼻音韻母】
    im_piau = im_piau.replace("ⁿ", "nn", 1)

    # 轉換音標中【韻母】為【o͘】（oo長音）的特殊處理
    im_piau = handle_o_dot(im_piau)

    # # 聲調符號對映調號數值的轉換字典
    # tiau_fu_mapping = {
    #     "\u0300": "3",   # 3 陰去: ò
    #     "\u0301": "2",   # 2 陰上: ó
    #     "\u0302": "5",   # 5 陽平: ô
    #     "\u0304": "7",   # 7 陽去: ō
    #     "\u0306": "9",   # 9 輕声: ő
    #     "\u030C": "6",   # 6 陽上: ǒ
    #     "\u030D": "8",   # 8 陽入: o̍
    # }
    # 轉換音標中【韻母】部份，不含【o͘】（oo長音）的特殊處理
    letters, tone = separate_tone(im_piau)   # 無調符音標：im_piau
    if tone:
        tiau_ho = tiau_fu_mapping[tone]
    else:
        tiau_ho = ""

    # 以【無調符音標】，轉換【韻母】
    sorted_keys = sorted(un_bu_mapping, key=len, reverse=True)
    for key in sorted_keys:
        if key in letters:
            letters = letters.replace(key, un_bu_mapping[key])
            break

    # 如若傳入之【音標】首字母為大寫，則將已轉成 "z" 或 "c" 之拼音字母改為大寫
    if su_ji.isupper():
        if letters[0] == "z":
            letters = "Z" + letters[1:]
        elif letters[0] == "c":
            letters = "C" + letters[1:]
        else:
            letters = su_ji + letters[1:]

    # 調符
    if tone:
        letters = apply_tone(letters, tone)

    return letters

def tng_tiau_ho(im_piau: str, kan_hua: bool = False) -> str:
    """
    將【帶調符音標】轉換為【帶調號音標】
    :param im_piau: str - 帶調符音標
    :param kan_hua: bool - 簡化：若是【簡化】，聲調值為 1 或 4 ，去除調號值
    :return: str - 帶調號音標
    """
    if im_piau == '': return ''  # noqa: E701
    # 遇標點符號，不做轉換處理，直接回傳
    if im_piau[-1] in PUNCTUATIONS:
        return im_piau

    # 若【音標】末端為數值，表音標已是【帶調號拼音】，直接回傳
    u_tiau_ho = True if im_piau[-1] in "123456789" else False
    if u_tiau_ho: return im_piau  # noqa: E701

    # 將傳入【音標】字串，以標準化之 NFC 組合格式，調整【帶調符拼音字母】；
    # 令以下之處理作業，不會發生【看似相同】的【帶調符拼音字母】，其實使用
    # 不同之 Unicode 編碼
    im_piau = unicodedata.normalize("NFC", im_piau)

    #--------------------------------------------------------------------------------
    # 以【元音及韻化輔音清單】，比對傳入之【音標】，找出對應之【基本拼音字母】與【調號】
    #--------------------------------------------------------------------------------
    tone_number = "1"  # 初始化調號為 1
    number = "1"  # 明確初始化 number 變數，以免未設定而發生錯誤
    for tone_mark, (base_char, number) in tiau_hu_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)
            break
    else:
        number = "1"  # 若沒有任何調符，number強制為1

    # 依是否要【簡化】之設定，處理【調號】1 或 4 是否要略去
    if kan_hua and number in ["1", "4"]:
        # 若是【簡化】，且聲調值為 1 或 4 ，去除調號值
        tone_number = ""
    else:
        # 若未要求【簡化】，聲調值置於【音標】末端
        if number not in ["1", "4"]:
            tone_number = number
        else:
            if im_piau[-1] in "hptk":
                # 【音標】末端為【hptk】之一，則為【陰入調】，聲調值為 4
                tone_number = "4"
            else:
                tone_number = "1"

    return im_piau + tone_number


# =========================================================================
# 將【漢字庫】查詢所得結果，解析出【台語音標】，並依據使用者設定輸出【漢字標音】
# =========================================================================
def format_han_ji_piau_im(value):
    if isinstance(value, str):
        return value  # 已是字串

    if isinstance(value, (list, tuple)):
        # return " ".join(filter(None, value))  # 序列 -> 去空值後 join
        return "".join(filter(None, value))  # 序列 -> 去空值後 join

    return str(value)  # 其他型別備援


def ca_ji_tng_piau_im(entry, han_ji_khoo: str, piau_im, piau_im_huat: str):
    """查字結果的單筆紀錄出標音：查詢【漢字庫】取得之【查找結果】，將之切分：聲、韻、調"""
    if han_ji_khoo == "河洛話":
        #-----------------------------------------------------------------
        # 【白話音】：依《河洛話漢字庫》標注【台語音標】和【方音符號】
        #-----------------------------------------------------------------
        # 將【台語音標】分解為【聲母】、【韻母】、【聲調】
        # siann_bu = entry['聲母']
        # un_bu = entry['韻母']
        # tiau_ho = entry['聲調']
        siann_bu = entry.get('聲母', '')
        un_bu = entry.get('韻母', '')
        tiau_ho = entry.get('聲調', '')
        un_bu = tai_gi_im_piau_tng_un_bu(un_bu)
        if tiau_ho == "6":
            # 若【聲調】為【6】，則將【聲調】改為【7】
            tiau_ho = "7"
    else:
        #-----------------------------------------------：------------------
        # 【文讀音】：依《廣韻字庫》標注【台語音標】和【方音符號】
        #-----------------------------------------------------------------
        siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(entry[0]['標音'])
        if siann_bu == "" or siann_bu is None:
            siann_bu = "ø"

    # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
    # tai_gi_im_piau = siann_bu + un_bu + tiau_ho
    tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])

    # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
    if (piau_im_huat == "十五音" or piau_im_huat == "雅俗通") and (siann_bu == "" or siann_bu is None):
        siann_bu = "ø"

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
        ok = False

    # 若 ok 為 False，表示轉換失敗，則將【台語音標】直接傳回
    if not ok:
        return tai_gi_im_piau, ""
    else:
        return tai_gi_im_piau, format_han_ji_piau_im(han_ji_piau_im)



def ca_ji_kiat_ko_tng_piau_im(result, han_ji_khoo: str, piau_im, piau_im_huat: str):
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
        if siann_bu == "" or siann_bu is None:
            siann_bu = "ø"

    # 將【聲母】、【韻母】、【聲調】，合併成【台語音標】
    # tai_gi_im_piau = siann_bu + un_bu + tiau_ho
    tai_gi_im_piau = ''.join([siann_bu, un_bu, tiau_ho])

    # 標音法為：【十五音】或【雅俗通】，且【聲母】為空值，則將【聲母】設為【ø】
    if (piau_im_huat == "十五音" or piau_im_huat == "雅俗通") and (siann_bu == "" or siann_bu is None):
        siann_bu = "ø"

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
        ok = False

    # 若 ok 為 False，表示轉換失敗，則將【台語音標】直接傳回
    if not ok:
        return tai_gi_im_piau, ""
    else:
        return tai_gi_im_piau, format_han_ji_piau_im(han_ji_piau_im)


# =========================================================================
# 將首字母為大寫之羅馬拼音字母轉換為小寫（只處理第一個字母）
# =========================================================================
def normalize_im_piau_case(im_piau: str) -> str:
    im_piau = unicodedata.normalize("NFC", im_piau)  # 先標準化 Unicode
    return im_piau[0].lower() + im_piau[1:] if im_piau else im_piau


# =========================================================================
# 台語音標 → 台羅拼音（TLPA → TL）轉換函數
# =========================================================================
def convert_tlpa_to_tl(tai_gi_im_piau: str) -> str:
    """
    轉換台語音標（TLPA）為台羅拼音（TL）。
    """
    if not tai_gi_im_piau:
        return ""

    # 第一次替換：c → tsh
    tai_lo_im_piau = re.sub(r'\bc', 'tsh', tai_gi_im_piau)
    # 第二次替換：z → ts，使用上一次替換之後的 tai_lo_im_piau
    tai_lo_im_piau = re.sub(r'\bz', 'ts', tai_lo_im_piau)

    return tai_lo_im_piau


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


# =========================================================================
# 台羅拼音 → 台語音標（TL → TLPA）轉換函數
# =========================================================================

# 聲調符號對應表（帶調號母音 → 對應數字）
# "a": ("a", "1"), "á": ("a", "2"), "à": ("a", "3"), "a": ("a", "4"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "a̍": ("a", "8"), "a̋": ("a", "9"),
# "e": ("e", "1"), "é": ("e", "2"), "è": ("e", "3"), "e": ("e", "4"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "e̍": ("e", "8"), "e̋": ("e", "9"),
# "i": ("i", "1"), "í": ("i", "2"), "ì": ("i", "3"), "i": ("i", "4"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "i̍": ("i", "8"), "i̋": ("i", "9"),
# "o": ("o", "1"), "ó": ("o", "2"), "ò": ("o", "3"), "o": ("o", "4"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "o̍": ("o", "8"), "ő ": ("o", "9"),
# "u": ("u", "1"), "ú": ("u", "2"), "ù": ("u", "3"), "u": ("u", "4"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "u̍": ("u", "8"), "ű ": ("u", "9"),
# "m": ("m", "1"), "ḿ": ("m", "2"), "m̀": ("m", "3"), "m": ("m", "4"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̍": ("m", "8"), "m̋": ("m", "9"),
# "n": ("n", "1"), "ń": ("n", "2"), "ǹ": ("n", "3"), "n": ("n", "4"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̍": ("n", "8"), "n̋": ("n", "9"),
tone_mapping = {
    "á": ("a", "2"), "à": ("a", "3"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "a̍": ("a", "8"), "a̋": ("a", "9"),
    "é": ("e", "2"), "è": ("e", "3"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "e̍": ("e", "8"), "e̋": ("e", "9"),
    "í": ("i", "2"), "ì": ("i", "3"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "i̍": ("i", "8"), "i̋": ("i", "9"),
    "ó": ("o", "2"), "ò": ("o", "3"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "o̍": ("o", "8"), "ő ": ("o", "9"),
    "ú": ("u", "2"), "ù": ("u", "3"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "u̍": ("u", "8"), "ű ": ("u", "9"),
    "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̍": ("m", "8"), "m̋": ("m", "9"),
    "ń": ("n", "2"), "ǹ": ("n", "3"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̍": ("n", "8"), "n̋": ("n", "9"),
}

# 聲母轉換規則（台羅拼音 → 台語音標+）
initials_mapping = {
    "tsh": "c",
    "ts": "z"
}


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

    # 3. 若無聲調符號，需判斷是陰平調（調號 1）或陰入調（調號 4）
    # 入聲調的羅馬字必須以 p, t, k, h 結尾
    if tai_lo:
        # 若最後一個字元已經是數字，則不需再加調號
        if tai_lo[-1].isdigit():
            return tai_lo

        last_char = tai_lo[-1].lower()
        if last_char in ['p', 't', 'k', 'h']:
            # 以 p/t/k/h 結尾的是陰入調（調號 4）
            return tai_lo + "4"
        else:
            # 其他情況是陰平調（調號 1）
            return tai_lo + "1"

    return tai_lo  # 若無內容則不變更


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


# =========================================================
# 判斷是否為標點符號的輔助函數
# =========================================================
def is_punctuation(char):
    """判斷是否為標點符號"""
    import unicodedata
    if not char or not isinstance(char, str):
        return False
    if len(char) != 1:
        return False  # ✅ 只允許單一字元
    return unicodedata.category(char)[0] in {'P', 'S'}

# =========================================================
# 想要僅針對漢字進行檢查，而不包括其他語言的字母，可用 Unicode 範圍來判斷。
# 漢字的 Unicode 範圍： [\u4e00-\u9fff] (包括中日韓越所有漢字)
# =========================================================
def char_is_han_ji(char):
    return '\u4e00' <= char <= '\u9fff'

def is_han_ji(char):
    """檢查字元是否為漢字"""
    if not isinstance(char, str) or len(char) != 1:
        return False
    return 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(char, '')

def is_valid_han_ji(char):
    if char is None:
        return False
    else:
        char = char.strip()

    punctuation_marks = "，。！？；：、（）「」『』《》……"
    return char not in punctuation_marks

def extract_han_ji(text):
    if not isinstance(text, str):
        text = str(text or "")
    return ''.join([c for c in text if is_han_ji(c)])

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
    return ''.join(superscript_digit_mapping.get(char, char) for char in input_str)


# ============================================================================
# 將使用【上標數字】表示的【調號】，轉換為普通數字
# ============================================================================
def split_tai_lo(input_str):
    # 將上標數字替換為普通數字
    input_str = replace_superscript_digits(input_str)
    # 使用正則表達式匹配聲母、韻母和調號
    # pattern = r'^([ptkhmnljw]?)([aeiouáéíóúâêîôûäëïöü]+)([0-9])?$'
    pattern = r'^([ptkhmnlzcsjw]?)([aeiouáéíóúâêîôûäëïöü]+)([0-9])?$'
    match = re.match(pattern, input_str)
    if match:
        siann_bu = match.group(1)
        un_bu = match.group(2)
        tiau_ho = match.group(3) if match.group(3) else '1'  # 默認調號為1
        return siann_bu, un_bu, tiau_ho
    else:
        return None, None, None

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


# ==========================================================
# 台語音標轉換為【漢字標音】之注音符號或羅馬字音標
# ==========================================================
def im_piau_iong_tiau_ho(im_piau: str) -> bool:
    """
    判斷台語音標/台羅拚音是否帶有調號（最後一個字元是否為數字）

    Args:
        tai_gi_im_piau: 台語音標字串

    Returns:
        True: 有調號
        False: 無調號
    """
    if not im_piau:
        return False

    return im_piau[-1].isdigit()


def tlpa_tng_han_ji_piau_im(piau_im, piau_im_huat, tai_gi_im_piau):
    # 若傳入函數之參數【tai_gi_im_piau】，為【帶聲調符號】之【台語音標】，
    # 則先轉換為【帶調號】之【台語音標】
    if not im_piau_iong_tiau_ho(tai_gi_im_piau):
        tai_gi_im_piau_iong_tiau_ho = convert_tl_with_tiau_hu_to_tlpa(tai_gi_im_piau)
    else:
        tai_gi_im_piau_iong_tiau_ho = tai_gi_im_piau
    siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_im_piau_iong_tiau_ho)

    if siann_bu == "" or siann_bu is None:
        # siann_bu = "Ø"
        siann_bu = 'ø'

    han_ji_piau_im = ""
    try:
        han_ji_piau_im = piau_im.han_ji_piau_im_tng_huan(
            piau_im_huat=piau_im_huat,
            siann_bu=siann_bu,
            un_bu=un_bu,
            tiau_ho=tiau_ho,
        )
        if not han_ji_piau_im: # 傳回：空字串、""、False、0，表示【漢字標音】轉換失敗
            logging_exc_error(msg=f"【台語音標】：[{tai_gi_im_piau}]，無法以【{piau_im_huat}】漢字標音法轉換！", error=None)
            han_ji_piau_im = ""
    except Exception as e:
        logging_exception(f"piau_im.han_ji_piau_im_tng_huan() 發生執行時期錯誤: 【台語音標】：{tai_gi_im_piau}", e)

    return han_ji_piau_im

# =========================================================
# 判斷是否為標點符號的輔助函數
# =========================================================
# def is_valid_han_ji(char):
#     return is_punctuation(char) or is_chinese_char(char)

# # 方音符號轉換為【台語音標】
# def hong_im_tng_tai_gi_im_piau(siann, un, tiau, cursor):
#     """
#     根據傳入的方音符號聲母、韻母、聲調，轉換成對應的台語音標
#     :param siann: 聲母 (方音符號)
#     :param un: 韻母 (方音符號)
#     :param tiau: 聲調 (方音符號)
#     :param cursor: 數據庫游標
#     :return: 包含台語音標的字典
#     """
#     # 查詢聲母表，將方音符號的聲母轉換成台語音標
#     cursor.execute("SELECT 台語音標 FROM 聲母對照表 WHERE 方音符號 = ?", (siann,))
#     siann_result = cursor.fetchone()
#     if siann_result:
#         tai_gi_siann = siann_result[0]  # 取得台語音標
#     else:
#         tai_gi_siann = ''  # 無聲母的情況

#     # 查詢韻母表，將方音符號的韻母轉換成台語音標
#     cursor.execute("SELECT 台語音標 FROM 韻母對照表 WHERE 方音符號 = ?", (un,))
#     un_result = cursor.fetchone()
#     if un_result:
#         tai_gi_un = un_result[0]  # 取得台語音標
#     else:
#         tai_gi_un = ''

#     # 查詢聲調表，將方音符號的聲調轉換成台語音標
#     # cursor.execute("SELECT 方音符號調符 FROM 聲調對照表 WHERE 台羅調號 = ?", (tiau,))
#     # tiau_result = cursor.fetchone()
#     # if tiau_result:
#     #     tai_gi_tiau = tiau_result[0]  # 取得台語音標
#     # else:
#     #     tai_gi_tiau = ''
#     tai_gi_tiau = tiau

#     return {
#         '台語音標': f"{tai_gi_siann}{tai_gi_un}{tai_gi_tiau}",
#         '聲母': tai_gi_siann,
#         '韻母': tai_gi_un,
#         '聲調': tai_gi_tiau,
#     }


# 台語音標轉換為方音符號
def TL_Tng_Zu_Im(siann_bu, un_bu, siann_tiau, cursor):
    """
    根據傳入的台語音標聲母、韻母、聲調，轉換成對應的方音符號
    :param siann_bu: 聲母 (台語音標)
    :param un_bu: 韻母 (台語音標)
    :param siann_tiau: 聲調 (台語音標中的數字)
    :param cursor: 數據庫游標
    :return: 包含方音符號的字典
    """

    # 查詢聲母表，將台語音標的聲母轉換成方音符號
    cursor.execute("SELECT 方音符號 FROM 聲母對照表 WHERE 台語音標 = ?", (siann_bu,))
    siann_bu_result = cursor.fetchone()
    if siann_bu_result:
        zu_im_siann_bu = siann_bu_result[0]  # 取得方音符號
    else:
        zu_im_siann_bu = ''  # 無聲母的情況

    # 查詢韻母表，將台語音標的韻母轉換成方音符號
    # cursor.execute("SELECT 方音符號 FROM 韻母表 WHERE 台語音標 = ?", (un_bu,))
    cursor.execute("SELECT 方音符號 FROM 韻母對照表 WHERE 台語音標 = ?", (un_bu,))
    un_bu_result = cursor.fetchone()
    if un_bu_result:
        zu_im_un_bu = un_bu_result[0]  # 取得方音符號
    else:
        zu_im_un_bu = ''

    # 查詢聲調表，將台語音標的聲調轉換成方音符號
    cursor.execute("SELECT 方音符號調符 FROM 聲調對照表 WHERE 台羅調號 = ?", (siann_tiau,))
    siann_tiau_result = cursor.fetchone()
    if siann_tiau_result:
        zu_im_siann_tiau = siann_tiau_result[0]  # 取得方音符號
    else:
        zu_im_siann_tiau = ''

    #=======================================================================
    # 【聲母】校調
    #
    # 齒間音【聲母】：ㄗ、ㄘ、ㄙ、ㆡ，若其後所接【韻母】之第一個符號亦為：ㄧ、ㆪ時，須變改
    # 為：ㄐ、ㄑ、ㄒ、ㆢ。
    #-----------------------------------------------------------------------
    # 參考 RIME 輸入法如下規則：
    # - xform/ㄗ(ㄧ|ㆪ)/ㄐ$1/
    # - xform/ㄘ(ㄧ|ㆪ)/ㄑ$1/
    # - xform/ㄙ(ㄧ|ㆪ)/ㄒ$1/
    # - xform/ㆡ(ㄧ|ㆪ)/ㆢ$1/
    #=======================================================================

    # 比對聲母是否為 ㄗ、ㄘ、ㄙ、ㆡ，且韻母的第一個符號是 ㄧ 或 ㆪ
    if siann_bu == 'z' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㄐ'
    elif siann_bu == 'c' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㄑ'
    elif siann_bu == 's' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㄒ'
    elif siann_bu == 'j' and (un_bu[0] == 'i' or un_bu == 'inn'):
        zu_im_siann_bu = 'ㆢ'

    return {
        '注音符號': f"{zu_im_siann_bu}{zu_im_un_bu}{zu_im_siann_tiau}",
        '聲母': zu_im_siann_bu,
        '韻母': zu_im_un_bu,
        '聲調': zu_im_siann_tiau
    }


def dict_to_str(zu_im_hu_ho):
    return f"{zu_im_hu_ho['聲母']}{zu_im_hu_ho['韻母']}{zu_im_hu_ho['聲調']}"


# ==========================================================

class PiauIm:

    TONE_MARKS = {
        "十五音": {
            1: "一",
            2: "二",
            3: "三",
            4: "四",
            5: "五",
            7: "七",
            8: "八"
        },
        "方音符號": {
            1: "",
            2: "ˋ",
            3: "˪",
            4: "",
            5: "ˊ",
            7: "˫",
            8: "\u02D9",
            0: "\u2070",
        },
        "閩拼方案": {
            1: "\u0304",
            2: "\u0341",
            3: "\u030C",
            5: "\u0300",
            6: "\u0302",
            7: "\u0304",
            8: "\u0341"
        },
        "台羅拼音": {
            1: "",
            2: "\u0301",
            3: "\u0300",
            4: "",
            5: "\u0302",
            6: "\u030C",
            7: "\u0304",
            8: "\u030D",
            9: "\u030B"
        }
    }

    Hong_Im_Tiau_Hu_Dict = {
        "ˋ"    : 2,
        "˪"     : 3,
        "ˊ"    : 5,
        "˫"     : 7,
        "\u02D9": 8,
        "\u2070": 0,
    }

    # def __init__(self, han_ji_khoo="漢語標音"):
    def __init__(self, han_ji_khoo="漢語標音", cursor=None):
        self.Siann_Bu_Dict = None
        self.Un_Bu_Dict = None
        self.cursor = cursor  # 將 cursor 存入物件屬性
        self.init_piau_im_dict(han_ji_khoo)
        self.TL_pattern1 = re.compile(r"(uai|uan|uah|ueh|ee|ei|oo)", re.I)
        self.TL_pattern2 = re.compile(r"(o|e|a|u|i|n|m)", re.I)
        self.POJ_pattern1 = re.compile(r"(oai|oan|oah|oeh|ee|ei)", re.I)
        self.POJ_pattern2 = re.compile(r"(o|e|a|u|i|n|m)", re.I)
        self.HongImTiauHu = re.compile(r"ˋ|˪|ˊ|˫|\u02D9", re.I)

    def set_cursor(self, cursor):
        """
        設定資料庫 cursor 物件
        :param cursor: 資料庫 cursor 物件
        """
        self.cursor = cursor

    def get_cursor(self):
        """
        取得資料庫 cursor 物件
        :return: cursor 物件
        """
        return self.cursor

    def _init_siann_bu_dict(self):
        """
        初始化聲母對照表，使用 cursor 進行 SQL 查詢
        """
        if not self.cursor:
            raise ValueError("資料庫 cursor 未設定，無法執行查詢")
        self.cursor.execute("SELECT * FROM 聲母對照表")
        rows = self.cursor.fetchall()
        #------------------------------------------------------------------
        # 從查詢結果中提取資料並將其整理成一個字典
        #------------------------------------------------------------------
        # siann_bu_dict = {row[1]: {'台語音標': row[1], '國際音標': row[2]} for row in rows}
        siann_bu_dict = {}          # 初始化字典
        for row in rows:
            siann_bu_dict[row[1]] = {
                '台語音標': row[1],
                '國際音標': row[2],
                '台羅拼音': row[3],
                '白話字':   row[4],
                '閩拼方案': row[5],
                '方音符號': row[6],
                '十五音':   row[7],
            }
        return siann_bu_dict

    def _init_un_bu_dict(self):
        """
        初始化韻母對照表，使用 cursor 進行 SQL 查詢
        """
        if not self.cursor:
            raise ValueError("資料庫 cursor 未設定，無法執行查詢")
        self.cursor.execute("SELECT * FROM 韻母對照表")
        rows = self.cursor.fetchall()
        #------------------------------------------------------------------
        # 設定【韻母對照表】用字典
        # un_bu_dict = {row[1]: {'台語音標': row[1], '國際音標': row[2]} for row in rows}
        #------------------------------------------------------------------
        # 初始化字典
        un_bu_dict = {}
        # 從查詢結果中提取資料並將其整理成一個字典
        for row in rows:
            un_bu_dict[row[1]] = {
                '台語音標': row[1],
                '國際音標': row[2],
                '台羅拼音': row[3],
                '白話字': row[4],
                '閩拼方案': row[5],
                '方音符號': row[6],
                '十五音': row[7],
                '十五音舒促聲': row[8],
                '十五音序': int(row[9]),
            }
        return un_bu_dict

    def init_piau_im_dict(self, han_ji_khoo):
        """
        初始化聲母與韻母字典
        :param han_ji_khoo: 標音類型
        """
        db_name = 'Ho_Lok_Ue.db' if han_ji_khoo == "河洛話" else 'Han_Ji_Piau_Im.db'
        with sqlite3.connect(db_name) as conn:
            self.cursor = conn.cursor() if not self.cursor else self.cursor
            self.Siann_Bu_Dict = self._init_siann_bu_dict()
            self.Un_Bu_Dict = self._init_un_bu_dict()

    #================================================================
    # 在韻母加調號：白話字(POJ)與台羅(TL)同
    #================================================================
    def un_bu_ga_tiau_ho(self, guan_im, tiau):
        tiau_hu_dict = {
            1: "",
            2: "\u0301",
            3: "\u0300",
            4: "",
            5: "\u0302",
            6: "\u030C",
            7: "\u0304",
            8: "\u030D",
            9: "\u030B",
        }
        guan_im_u_ga_tiau_ho = f"{guan_im}{tiau_hu_dict[int(tiau)]}"
        return guan_im_u_ga_tiau_ho

    #================================================================
    # 在韻母加調號：閩拼方案(BP)
    #================================================================
    def bp_un_bu_ga_tiau_ho(self, guan_im, tiau):
        tiau_hu_dict = {
            1: "\u0304",  # 陰平
            2: "\u0341",  # 陽平
            3: "\u030C",  # 上声
            5: "\u0300",  # 陰去
            6: "\u0302",  # 陽去
            7: "\u0304",  # 陰入
            8: "\u0341",  # 陽入
        }
        return f"{guan_im}{tiau_hu_dict[tiau]}"

    #================================================================
    # 台羅拼音（TL）
    # 順序：《o＞e＞a＞u＞i＞ng＞m》；而 ng 標示在字母 n 上。
    #================================================================
    def TL_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "台羅拼音"

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        if siann_bu == "" or siann_bu is None or siann_bu == "Ø" or siann_bu == "ø":
            siann = ""
        else:
            siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]

        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        piau_im = f"{siann}{un}"

        # 韻母為複元音
        searchObj = self.TL_pattern1.search(piau_im)
        if searchObj:
            found = searchObj.group(1)
            un_chars = list(found)
            idx = 0
            if found == "ee" or found == "ei" or found == "oo":
                idx = 0
            else:
                # found = uai/uan/uah/ueh
                idx = 1
            guan_im = un_chars[idx]
            un_chars[idx] = self.un_bu_ga_tiau_ho(guan_im, tiau_ho)
            un_str = "".join(un_chars)
            piau_im = piau_im.replace(found, un_str)
        else:
            # 韻母為單元音或鼻音韻
            searchObj2 = self.TL_pattern2.search(piau_im)
            if searchObj2:
                found = searchObj2.group(1)
                guan_im = found
                new_un = self.un_bu_ga_tiau_ho(guan_im, tiau_ho)
                piau_im = piau_im.replace(found, new_un)

        return piau_im

    #================================================================
    # 白話字（POJ）
    # 順序：《o＞e＞a＞u＞i＞ng＞m》；而 ng 標示在字母 n 上。
    # 例外：
    #  - oai、oan、oat、oah 標在 a 上。
    #  - oeh 標在 e 上。
    #================================================================
    def POJ_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "白話字"

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        if siann_bu == "" or siann_bu is None or siann_bu == "Ø" or siann_bu == "ø":
            siann = ""
        else:
            siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]

        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        piau_im = f"{siann}{un}"

        # 韻母為複元音
        searchObj = self.POJ_pattern1.search(piau_im)
        if searchObj:
            found = searchObj.group(1)
            un_chars = list(found)
            idx = 0
            if found == "ee" or found == "ei":
                idx = 0
            else:
                # found = oai/oan/oah/oeh
                idx = 1
            guan_im = un_chars[idx]
            un_chars[idx] = self.un_bu_ga_tiau_ho(guan_im, tiau_ho)
            un_str = "".join(un_chars)
            piau_im = piau_im.replace(found, un_str)
        else:
            # 韻母為單元音或鼻音韻
            searchObj2 = self.POJ_pattern2.search(piau_im)
            if searchObj2:
                found = searchObj2.group(1)
                guan_im = found
                new_un = self.un_bu_ga_tiau_ho(guan_im, tiau_ho)
                piau_im = piau_im.replace(found, new_un)

        return piau_im


    #================================================================
    # 閩拼（BP）
    #================================================================
    #
    # 《閩拼方案規範》
    #
    # 【零聲母，音節開頭為母音 'i' 或 'u' 之轉換規則】
    # 音節為母音[i]或[u]開頭時(零聲母)，會視情況將其母音改寫或增加一個y或w字首，此規則與漢語拼音方案相同。
    #   （1）傳入之 siann_bu 為："Ø"/""/None ；且 un_bu 的第一個羅馬拼音字母為「i」時：
    #       (a) i 後無其它【韻母】，則在 "i" 之前，增添「y」。
    #           【例】： 【依】= "" + "i" + "1" → un_bu = "yi"。
    #                   【因】= "" + "in" + "1" →  un_bu = "yin"。
    #       (b) i 後有其它羅馬字母，則將 "i" 改為「y」。
    #           【例】：【鴉】= "" + "ia" + "1" → un_bu = "ya"。
    #                   【煙】 = "" + "ian" + "1" → un_bu = "yan"。
    #   （2）傳入之 siann_bu 為："Ø"/""/None ；且 un_bu 的第一個羅馬拼音字母為「u」時：
    #       (a) u 後無其它【韻母】，則在 "u" 之前，增添「w」。
    #           【例】： 【烏】= "" + "u" + "1" → un_bu = "wu"。
    #                   【溫】= "" + "un" + "1" →  un_bu = "wun"。
    #       (b) u 後有其它羅馬字母，則將 "u" 改為「w」。
    #           【例】：【蛙】= "" + "ua" + "1" → un_bu = "wa"。
    #                   【彎】 = "" + "uan" + "1" → un_bu = "wan"。
    #
    # 【調號標示規則】
    # 當一個音節有多個字母時，調號得標示在響度最大的字母上面（通常在韻腹）。由規則可以判定確切的字母：
    #
    #  - 響度優先順序： a > oo > (e = o) > (i = u)〈低元音 > 高元音 > 無擦通音 > 擦音 > 塞音〉
    #  - 二合字母 iu 及 ui ，調號都標在後一個字母上；因為前一個字母是介音。
    #  - m 作韻腹時則標於字母 m 上。
    #  - 二合字母 oo 及 ng，標於前一個字母上；比如 ng 標示在字母 n 上。
    #  - 三合字母 ere，標於最後的字母 e 上。
    #================================================================
    def _get_BP_syllable(self, siann_bu, un_bu, tiau_ho, with_tone_number=True) -> str:
        """
        產生未附聲調符號的【閩拼音節】，可選擇附數字調號或不附（方便後續加符號）
        """
        piau_im_huat = "閩拼方案"
        # 將「台羅八聲調」轉換成閩拼使用的調號
        Tiau_Ho_Remap = {
            0: 0,  # 輕聲: 40
            1: 1,  # 陰平: 44
            2: 3,  # 上聲：53
            3: 5,  # 陰去：21
            4: 7,  # 上聲：53
            5: 2,  # 陽平：24
            7: 6,  # 陰入：3?
            8: 8,  # 陽入：4?
        }

        # 將表【調號】之【上標數值】字母轉換為標準字母
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        try:
            # 【調號】值為 6 時，轉換成：7
            tiau_ho_int = 7 if int(tiau_ho) == 6 else int(tiau_ho)
        except (TypeError, ValueError):
            logging_warning(f"無法將【調號】轉為整數: {tiau_ho}")
            return ""  # 避免程式執行至此，抛出執行時期錯誤，終止程式執行

        tiau = Tiau_Ho_Remap.get(tiau_ho_int)
        if tiau is None:
            logging_warning(f"無法對映【調號】: {tiau_ho_int}；\n【聲母】: {siann_bu}；【韻母】: {un_bu}；【調號】: {tiau_ho}")
            return ""  # 避免程式執行至此，抛出執行時期錯誤，終止程式執行

        # 聲母轉換
        if siann_bu in ("", None, "Ø", "ø"):
            # 遇有零聲母的特殊處理
            siann = ""
        else:
            siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]
            if not siann:
                logging_warning(f"無法對映【聲母】: {siann_bu}；\n【聲母】: {siann_bu}；【韻母】: {un_bu}；【調號】: {tiau_ho}")
                return ""  # 避免程式執行至此，抛出執行時期錯誤，終止程式執行

        # 韻母轉換
        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        if not un:
            logging_warning(f"無法對映【韻母】: {un_bu}；\n【聲母】: {siann_bu}；【韻母】: {un_bu}；【調號】: {tiau_ho}")
            return ""  # 避免程式執行至此，抛出執行時期錯誤，終止程式執行

        # 閩拼特例：零聲母 + i/u 開頭的韻母
        # 當聲母為「空白」，韻母為首之【羅馬拼音字母】為：i 或 u 時之調整作業
        if siann == "":
            if un.startswith("i"):
                # un = "y" + un[1:] if len(un) > 1 else "y" + un
                if len(un) == 1:
                    un = "y" + un  # i 後無其它韻母字母，增添 y
                else:
                    un = "y" + un[1:]  # i 後有其它韻母字母，將 i 改為 y
            elif un.startswith("u"):
                # un = "w" + un[1:] if len(un) > 1 else "w" + un
                if len(un) == 1:
                    un = "w" + un  # u 後無其它韻母字母，增添 w
                else:
                    un = "w" + un[1:]  # u 後有其它韻母字母，將 u 改為 w

        syllable = (siann, un, str(tiau)) if with_tone_number else (siann, un, "")
        return syllable

    def BP_piau_im(self, siann_bu, un_bu, tiau_ho) -> str :
        """將傳入的「TLPA+音標」之【聲母】、【韻母】、【調號】轉換為【閩拼方案】的音標。
        若是執行過程中發生錯誤，則回傳空字串。

        Args:
            siann_bu (str): TLPA+音標的聲母
            un_bu (str): TLPA+音標的韻母
            tiau_ho (str): TLPA+音標的調號

        Returns:
            str: 閩拼方案的音標
        """
        result: Optional[str] = self._get_BP_syllable(
            siann_bu, un_bu, tiau_ho, with_tone_number=True
        )
        if result is None:
            # 已由 _get_BP_syllable() 紀錄 warning；此處只需回傳預設值
            return ""

        bp_im_piau = f"{result[0]}{result[1]}{result[2]}"
        return bp_im_piau

    def BP_piau_im_with_tiau_hu(self, siann_bu, un_bu, tiau_ho):
        #----------------------------------------
        # 將【台語音標】轉換成 【閩拼音標】
        #----------------------------------------
        # siann, un, tiau = self._get_BP_syllable(siann_bu, un_bu, tiau_ho)

        # tlpa_im_piau = f"{siann_bu}{un_bu}{tiau_ho}"
        # bp_im_piau_list = convert_TLPA_to_BP(tlpa_im_piau)
        # bp_im_piau = ''.join(bp_im_piau_list) if bp_im_piau_list[0] is not None else None

        # 零聲母處理
        if siann_bu == "" or siann_bu is None or siann_bu == "ø":
            tlpa_im_piau = f"{un_bu}{tiau_ho}"
        else:
            tlpa_im_piau = f"{siann_bu}{un_bu}{tiau_ho}"

        #----------------------------------------
        # 將【閩拼音標】之【韻母】加上【聲調】符號
        #----------------------------------------
        bp_with_tone = convert_TLPA_to_BP_with_tone_marks(tlpa_im_piau)

        return bp_with_tone


    # def BP_piau_im_with_tiau_hu(self, siann_bu, un_bu, tiau_ho):
    #     #----------------------------------------
    #     # 取得【閩拼方案】之【聲母】、【韻母】、【調號】
    #     #----------------------------------------
    #     # siann, un, tiau = self._get_BP_syllable(siann_bu, un_bu, tiau_ho, with_tone_number=False)
    #     siann, un, tiau = self._get_BP_syllable(siann_bu, un_bu, tiau_ho)

    #     #----------------------------------------
    #     # 在【韻母】之【羅馬拼音字母】之上標示【聲調】符號
    #     #----------------------------------------

    #     # 韻腹標調位置規則
    #     pattern = r"(a|oo|ere|iu|ui|ng|e|o|i|u|m)"
    #     match = re.search(pattern, un, re.I)
    #     if match:
    #         found = match.group(1)
    #         idx = {"iu": 1, "ui": 1, "oo": 0, "ng": 0, "ere": 2}.get(found, 0)
    #         target = list(found)
    #         target[idx] = self.bp_un_bu_ga_tiau_ho(target[idx], tiau)
    #         un = un.replace(found, ''.join(target))

    #     return f"{siann}{un}"

    #================================================================
    # 閩拚注音（BP Zu Im）
    #================================================================
    def BP_zu_im(self, siann_bu, un_bu, tiau_ho):
        # piau_im_huat = "閩拚注音"

        # 將上標數字替換為普通數字
        # tiau_ho = replace_superscript_digits(str(tiau_ho))

        # 將【台語音標】轉換成【閩拼音標】
        if siann_bu == 'ø' or siann_bu == 'Ø' or siann_bu == '':
            tlpa_im_piau = f"{un_bu}{tiau_ho}"
        else:
            tlpa_im_piau = f"{siann_bu}{un_bu}{tiau_ho}"
        bp_siann, bp_un, bp_tiau = convert_TLPA_to_BP(tlpa_im_piau)
        # 將【閩拚音標】轉換成【閩拚注音】
        zu_im_siann, zu_im_un, zu_im_tiau = convert_bp_siann_un_tiau_to_zu_im(bp_siann, bp_un, bp_tiau)
        return f"{zu_im_siann}{zu_im_un}{zu_im_tiau}"

        # 將【閩拚音標】轉換成【閩拚注音】，再回傳【閩拚注音】之：聲母、韻母、調號
        # return convert_bp_siann_un_tiau_to_zu_im(bp_siann, bp_un, bp_tiau)


    #================================================================
    # 方音符號注音（TPS）
    #================================================================
    def TPS_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "方音符號"
        TPS_piau_im_remap_dict = {
            "ㄗㄧ": "ㄐㄧ",
            "ㄘㄧ": "ㄑㄧ",
            "ㄙㄧ": "ㄒㄧ",
            "ㆡㄧ": "ㆢㄧ",
        }
        Tiau_Ho_Remap = {
            1: "",
            2: "ˋ",
            3: "˪",
            4: "",
            5: "ˊ",
            7: "˫",
            8: "\u02D9",
            0: "\u02D9",
        }

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        # siann_bu = 'ø' if siann_bu == 'Ø' else siann_bu
        if siann_bu == 'Ø' or siann_bu == '':
            siann_bu = 'ø'
        siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]
        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        # tiau = self.TONE_MARKS[piau_im_huat][tiau_ho]
        tiau = Tiau_Ho_Remap[tiau_ho]
        piau_im = f"{siann}{un}{tiau}"

        pattern = r"(ㄗㄧ|ㄘㄧ|ㄙㄧ|ㆡㄧ)"
        searchObj = re.search(pattern, piau_im, re.M | re.I)
        if searchObj:
            key_value = searchObj.group(1)
            piau_im = piau_im.replace(key_value, TPS_piau_im_remap_dict[key_value])

        return piau_im

    #================================================================
    # 【國語注音符號第二式】簡稱為：注音二式，英文縮寫為：MPS2。
    #================================================================
    def MPS2_piau_im(self, siann_bu: str, un_bu: str, tiau_ho: str | int) -> str:
        """將傳入的「TLPA+音標」之【聲母】、【韻母】、【調號】轉換為【注音二式/MPS2】音標
        例：siann_bu='zi', un_bu='ann', tiau_ho=1  -> 'ziann1'

        Args:
            siann_bu (str): TLPA+音標的聲母
            un_bu (str): TLPA+音標的韻母
            tiau_ho (str): TLPA+音標的調號

        Returns:
            str: 閩拼方案的音標
        """
        # if siann_bu in ("", None, "Ø", "ø"):
        if siann_bu in ("Ø", "ø"):
            siann = ""
        else:
            siann = (siann_bu or "").strip().lower()
        un = (un_bu or "").strip().lower()
        tiau = str(tiau_ho).strip()

        # 若調號不是純數字（或空），可視需要過濾非數字
        if tiau and not tiau.isdigit():
            tiau = re.sub(r"\D+", "", tiau)

        TLPA_im_piau = f"{siann}{un}{tiau}"
        im_piau = convert_TLPA_to_MPS2(TLPA_im_piau)
        return im_piau
        # return convert_TLPA_to_MPS2(TLPA_im_piau)

    def Cu_Hong_Im_Hu_Ho_Tiau_Hu(self, tai_lo_tiau_ho):
        """
        取方音符號：將【台羅調號】轉換成【方音符號調號】
        :param tai_lo_tiau_ho: 台羅調號
        :return: 對應的方音符號調號
        """
        方音符號調號 = {
            1: '',
            2: 'ˋ',
            3: '˪',
            4: '',
            5: 'ˊ',
            6: 'ˋ',
            7: '˫',
            8: '˙'
        }
        return 方音符號調號.get(tai_lo_tiau_ho, None)

    #================================================================
    # 雅俗通十五音(Nga-Siok-Thong)
    #================================================================
    def NST_piau_im(self, siann_bu, un_bu, tiau_ho):
        # piau_im_huat = "雅俗通"
        # Tiau_Ho_Remap = {
        #     1: "上平",
        #     2: "上上",
        #     3: "上去",
        #     4: "上入",
        #     5: "下平",
        #     6: "下上",
        #     7: "下去",
        #     8: "下入",
        # }
        piau_im_huat = "十五音"
        Tiau_Ho_Remap = {
            1: "一",
            2: "二",
            3: "三",
            4: "四",
            5: "五",
            7: "七",
            8: "八",
        }

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        if siann_bu == "" or siann_bu is None or siann_bu == "Ø" or siann_bu == 'ø':
            siann = "英"
        else:
            siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]
        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        # tiau = self.TONE_MARKS[piau_im_huat][int(tiau_ho)]
        tiau = Tiau_Ho_Remap[tiau_ho]
        piau_im = f"{un}{tiau}{siann}"
        return piau_im

    #================================================================
    # 十五音(SNI:Sip-Ngoo-Im)
    #================================================================
    def SNI_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "十五音"
        Tiau_Ho_Remap = {
            1: "一",
            2: "二",
            3: "三",
            4: "四",
            5: "五",
            7: "七",
            8: "八",
        }

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        if siann_bu == "" or siann_bu is None or siann_bu == "Ø" or siann_bu == "ø":
            siann = "英"
        else:
            siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]
        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        # tiau = self.TONE_MARKS[piau_im_huat][int(tiau_ho)]
        tiau = Tiau_Ho_Remap[tiau_ho]
        piau_im = f"{siann}{un}{tiau}"
        return piau_im

    #================================================================
    # 轉換【漢字標音】
    #================================================================
    def han_ji_piau_im_tng_huan(self, piau_im_huat: str, siann_bu: str, un_bu: str, tiau_ho: str) -> str:
        """選擇並執行對應的注音方法"""
        if piau_im_huat == "十五音":
            return self.SNI_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "方音符號":
            return self.TPS_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "注音二式":
            return self.MPS2_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "雅俗通":
            return self.NST_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "白話字":
            return self.POJ_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "台羅拼音":
            return self.TL_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "閩拼調號":
            return self.BP_piau_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "閩拼調符":
            return self.BP_piau_im_with_tiau_hu(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "閩拼注音":
            return self.BP_zu_im(siann_bu, un_bu, tiau_ho)
        elif piau_im_huat == "台語音標":
            if siann_bu in ("", None, "Ø", "ø"):
                siann = ""
            else:
                siann = self.Siann_Bu_Dict[siann_bu]["台語音標"] or ""
            un = self.Un_Bu_Dict[un_bu]["台語音標"]
            return f"{siann}{un}{tiau_ho}"
        return ""


    def hong_im_tng_tai_gi_im_piau(self, siann, un, tiau):
        """
        將【方音符號】轉換為【台語音標】
        :param siann: 聲母 (方音符號)
        :param un: 韻母 (方音符號)
        :param tiau: 聲調 (方音符號)
        :return: (聲母, 韻母, 聲調) 的 tuple
        """
        if not self.cursor:
            raise ValueError("資料庫 cursor 未設定，無法執行查詢")

        # 查詢【聲母對照表】轉換【聲母】
        self.cursor.execute("SELECT 台語音標 FROM 聲母對照表 WHERE 方音符號 = ?", (siann,))
        siann_result = self.cursor.fetchone()
        tai_gi_siann = siann_result[0] if siann_result else ""

        # 查詢【韻母對照表】轉換【韻母】
        self.cursor.execute("SELECT 台語音標 FROM 韻母對照表 WHERE 方音符號 = ?", (un,))
        un_result = self.cursor.fetchone()
        tai_gi_un = un_result[0] if un_result else ""

        # 聲調不變，直接回傳
        tai_gi_tiau = tiau

        # return tai_gi_siann, tai_gi_un, tai_gi_tiau
        return {
            '台語音標': f"{tai_gi_siann}{tai_gi_un}{tai_gi_tiau}",
            '聲母': tai_gi_siann,
            '韻母': tai_gi_un,
            '聲調': tai_gi_tiau,
        }



def ut001():
    """測試：台語音標轉換韻母"""
    tai_gi_im_piau = 'kere1'
    tsing_khak_kiat_ko = 'kue1'
    print(f'轉換漢字【雞】的【漢字標音】：{tai_gi_im_piau}')
    result = tai_gi_im_piau_tng_un_bu(tai_gi_im_piau)
    print(f'轉換後應為：{tsing_khak_kiat_ko}')
    print(f'實際結果為：{result}')
    if result == tsing_khak_kiat_ko:
        print('測試成功')
    else:
        print('測試失敗')


def ut002():
    """測試：韻母之台語音標轉換"""
    un_bu = 'ir'
    tsing_khak_un_bu = un_bu_tng_huan_map_dict[un_bu]
    print(f'韻母【{un_bu}】')
    result = un_bu_tng_huan(un_bu)
    print(f'轉換後應為：{tsing_khak_un_bu}')
    print(f'實際結果為：{result}')
    if result == tsing_khak_un_bu:
        print('測試成功')
    else:
        print('測試失敗')


def ut003():
    """測試：帶聲調符號之台羅拼音轉換成聲調以數字表示之台羅拼音"""
    # 測試
    test_cases = ["lio̍k", "tāi", "bô", "siâu", "lâi", "pò", "tshi̍t", "tsuan", "giâm", "ló"]
    converted = [convert_tl_with_tiau_hu_to_tlpa(word) for word in test_cases]

    # 顯示轉換結果
    for original, converted_word in zip(test_cases, converted):
        print(f"{original} → {converted_word}")


def ut004():
    """測試：帶聲調符號之台羅拼音轉換成【台語音標】"""
    # 測試
    test_cases = ["lio̍k", "tāi", "bô", "siâu", "lâi", "pò", "tshi̍t", "tsuan", "giâm", "ló"]
    tai_lo_im_piau = [convert_tl_with_tiau_hu_to_tlpa(word) for word in test_cases]
    converted = [convert_tl_to_tlpa(word) for word in tai_lo_im_piau]

    # 顯示轉換結果
    for original, converted_word in zip(test_cases, converted):
        tai_gi_im_piau = "".join(converted_word)
        print(f"{original} → {tai_gi_im_piau}: {converted_word}")


def ut005():
    test_cases = ["Lio̍k", "Tshiâu", "Gua̍n", "Tsian"]
    converted = [normalize_im_piau_case(word) for word in test_cases]

    for original, converted_word in zip(test_cases, converted):
        print(f"{original} → {converted_word}")


if __name__ == "__main__":
    # # 測試：將【雞】kere1 轉換為【kue1】
    # print('==================================================================')
    # ut001()
    # # 測試：將【ir】轉換為【kue1】
    # print('==================================================================')
    # ut002()
    # # 測試：帶聲調符號之台羅拼音轉換成聲調以數字表示之台羅拼音
    # print('==================================================================')
    # ut003()
    # # 測試：帶聲調符號之台羅拼音轉換成【台語音標】
    # print('==================================================================')
    # ut004()
    # 測試：將首字母大寫的音標轉換成小寫
    print('==================================================================')
    ut005()
