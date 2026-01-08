# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import os
import re
import sys
import unicodedata

import xlwings as xw

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
# 程式區域函式
# =========================================================================

# 用途：從純文字檔案讀取資料並回傳 [漢字, ...] 之格式
def read_text_with_han_ji(filename: str = "tmp_p1_han_ji.txt") -> list:
    text_with_han_ji = []
    with open(filename, "r", encoding="utf-8") as f:
        # 先移除 `\u200b`，確保不會影響 TLPA 拼音對應
        lines = [re.sub(r"[\u200b]", "", line.strip()) for line in f if line.strip()]

    for i in range(0, len(lines), 1):
        han_ji = lines[i]
        text_with_han_ji.append(han_ji)

    return text_with_han_ji


# 用途：從純文字檔案讀取資料並回傳 [Im-Piau, ...] 之格式
def read_text_with_im_piau(filename: str = "ping_im.txt") -> list:
    text_with_tlpa = []
    with open(filename, "r", encoding="utf-8") as f:
        # 先移除 `\u200b`，確保不會影響 TLPA 拼音對應
        lines = [re.sub(r"[\u200b]", "", line.strip()) for line in f if line.strip()]

    # for i in range(0, len(lines), 2):
    for i in range(0, len(lines), 1):
        im_piau = lines[i].replace("-", " ")  # 替換 "-" 為空白字元
        text_with_tlpa.append((im_piau))

    return text_with_tlpa

def fix_im_piau_spacing(im_piau: str) -> str:
    """
    若音標的首字為漢字，則在首個漢字之後插入空白字元。
    如：「僇lâng」→「僇 lâng」
    """
    if im_piau and is_han_ji(im_piau[0]):
        return im_piau[0] + ' ' + im_piau[1:]
    return im_piau

def tu_bo_iong_ji_bu(ku: str) -> str:
    """
    清無用字母：清除控制字元
    :param ku: str - 句子輸入
    :return: str - 句子輸出
    """
    ku_clean = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2060-\u206f]', '', ku)
    return ku_clean

def cing_bo_iong_ji_bu(text: str) -> str:
    """_summary_
    清無用字母：清除控制字元
    Args:
        text (str): _description_

    Returns:
        str: _description_
    """
    return ''.join(
        ch for ch in text
        if unicodedata.category(ch)[0] != 'C'  # 排除所有類別為 Control (C) 的字元
    )

def normalize_sentence(im_piau_list):
    sentence = ''
    punctuations = {'.', ',', '!', '?', ':', ';'}

    for idx, item in enumerate(im_piau_list):
        if item in punctuations:
            sentence = sentence.rstrip() + item + ' '
        else:
            sentence += item + ' '

    sentence = sentence.strip()
    sentence = re.sub(r'\s+([.,!?;:])', r'\1', sentence)

    return sentence

# =========================================================================
# 讀取帶有【音標】標注【漢字】讀音之【純文字檔案】
# 回傳格式： [(漢字, TLPA), ...] 之格式
# =========================================================================

def read_text_with_tlpa(filename):
    text_with_tlpa = []
    with open(filename, "r", encoding="utf-8") as f:
        # 先移除 `\u200b`，確保不會影響 TLPA 拼音對應
        lines = [re.sub(r"[\u200b]", "", line.strip()) for line in f if line.strip() and not line.startswith("zh.wikipedia.org")]

    for i in range(0, len(lines), 2):
        hanzi = lines[i]
        tlpa = lines[i + 1].replace("-", " ")  # 替換 "-" 為空白字元
        text_with_tlpa.append((hanzi, tlpa))

    return text_with_tlpa

# =========================================================
# 音標整埋工具庫
# =========================================================
def is_im_piau(im_piau: str) -> bool:
    return im_piau in PUNCTUATIONS

# 用途：檢查是否為漢字
def is_han_ji(char):
    if not isinstance(char, str) or len(char) != 1:
        return False
    return 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(char, '')

# 清除控制字元：將 Unicode 中所有類別為 Control (C) 的字元移除
def cing_tu_khong_ze_ji_guan(text: str) -> str:
    """_summary_
    清除控制字元：將 Unicode 中所有類別為 Control (C) 的字元移除
    Args:
        text (str): _description_

    Returns:
        str: _description_
    """
    return ''.join(
        ch for ch in text
        if unicodedata.category(ch)[0] != 'C'  # 排除所有類別為 Control (C) 的字元
    )

def zing_li_zuan_ku(ku: str) -> str:
    """
    整理全句：移除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
    :param ku: str - 句子輸入
    :return: list - 斷詞結果
    """
    # 清除控制字元
    ku = cing_tu_khong_ze_ji_guan(ku)
    # 將 "-" 轉換成空白
    ku = ku.replace("-", " ")

    # 將標點符號前後加上空白
    ku = re.sub(f"([{''.join(re.escape(p) for p in PUNCTUATIONS)}])", r" \1 ", ku)

    # 移除多餘空白
    ku = re.sub(r"\s+", " ", ku).strip()

    return ku

def replace_superscript_digits(input_str: str) -> str:
    """將上標格式之數值字串轉換為一般數值字串

    Args:
        input_str (str): 上標數值字串

    Returns:
        str: 一般數值字串
    """
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


# =========================================================================
# 將【帶調符音標】轉換成【帶調號TLPA音標】
# =========================================================================

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

# 韻母轉換字典
un_bu_mapping = {
    'ee': 'e', 'ei': 'e', 'er': 'e', 'erh': 'eh', 'or': 'o', 'ere': 'ue', 'ereh': 'ueh',
    'ir': 'i', 'eng': 'ing', 'ek': 'ik', 'oa': 'ua', 'oe': 'ue', 'oai': 'uai',
    'ou': 'oo', 'onn': 'oonn', 'uei': 'ue', 'ueinn': 'uenn', 'ur': 'u',
}

# 聲調符號對映調號數值的轉換字典
tiau_fu_mapping = {
    "\u0300": "3",   # 3 陰去: ò
    "\u0301": "2",   # 2 陰上: ó
    "\u0302": "5",   # 5 陽平: ô
    "\u0304": "7",   # 7 陽去: ō
    "\u0306": "9",   # 9 輕声: ő
    "\u030C": "6",   # 6 陽上: ǒ
    "\u030D": "8",   # 8 陽入: o̍
}

# 調號與調符對映轉換字典
tiau_ho_mapping = {
    "3": "\u0300",   # 3 陰去: ò
    "2": "\u0301",   # 2 陰上: ó
    "5": "\u0302",   # 5 陽平: ô
    "7": "\u0304",   # 7 陽去: ō
    "9": "\u030B",   # 9 輕声: ő
    "6": "\u030C",   # 6 陽上: ǒ
    "8": "\u030D",   # 8 陽入: o̍
}


# 清理音標：整理音標中的字元組合，只留【拼音字母】，清除：標點符號、控制字元
def clean_im_piau(im_piau: str) -> str:
    # 移除標點符號
    im_piau = ''.join(ji_bu for ji_bu in im_piau if ji_bu not in PUNCTUATIONS)
    # 重新組合聲調符號（標準組合 NFC）
    im_piau = unicodedata.normalize("NFC", im_piau)
    return im_piau

# ---------------------------------------------------------
# 韻母轉換
# ---------------------------------------------------------

def separate_tone(im_piau):
    """拆解帶調字母為無調字母與調號"""
    decomposed = unicodedata.normalize('NFD', im_piau)
    letters = ''.join(c for c in decomposed if unicodedata.category(c) != 'Mn')
    tones = ''.join(c for c in decomposed if unicodedata.category(c) == 'Mn' and c != '\u0358')
    return letters, tones

def apply_tone(im_piau, tone):
    """
    根據 TLPA 響度優先規則，將聲調符號 tone 正確標示在母音上。
    """
    tone_priority = ['a', 'oo', 'e', 'o', 'i', 'u', 'm']
    lower = im_piau.lower()

    # 特例：ere → 最後 e
    if 'ere' in lower:
        idx = lower.rindex('e')
        return unicodedata.normalize('NFC', im_piau[:idx+1] + tone + im_piau[idx+1:])

    # ✅ 特例：iu / ui 雙母音（不限定結尾）
    for i in range(len(lower) - 1):
        if lower[i] == 'i' and lower[i+1] == 'u':
            return unicodedata.normalize('NFC', im_piau[:i+2] + tone + im_piau[i+2:])
        if lower[i] == 'u' and lower[i+1] == 'i':
            return unicodedata.normalize('NFC', im_piau[:i+2] + tone + im_piau[i+2:])

    # 特例：oo → 第一個 o
    if 'oo' in lower:
        idx = lower.index('oo')
        return unicodedata.normalize('NFC', im_piau[:idx+1] + tone + im_piau[idx+1:])

    # 響度優先分析
    best_idx = -1
    best_priority = len(tone_priority) + 1

    i = 0
    while i < len(lower):
        if lower[i] == 'o' and i+1 < len(lower) and lower[i+1] == 'o':
            current = 'oo'
            idx = i
            i += 1
        else:
            current = lower[i]
            idx = i

        if current in tone_priority:
            pri = tone_priority.index(current)
            if pri < best_priority:
                best_idx = idx
                best_priority = pri
            elif pri == best_priority and current in ['i', 'u']:
                best_idx = idx
        i += 1

    if best_idx != -1:
        return unicodedata.normalize('NFC', im_piau[:best_idx+1] + tone + im_piau[best_idx+1:])

    # fallback
    return unicodedata.normalize('NFC', im_piau[0] + tone + im_piau[1:])

# 處理 o͘ 韻母特殊情況的函數
def handle_o_dot(im_piau):
    # 依 Unicode 解構標準（NFD）分解傳入之【音標】，取得解構後之【拼音字母與調符】
    decomposed = unicodedata.normalize('NFD', im_piau)
    # 找出 o + 聲調 + 鼻化符號的特殊組合
    match = re.search(r'(o)([\u0300\u0301\u0302\u0304\u030B\u030C\u030D]?)(\u0358)', decomposed, re.I)
    if match:
        # 捕獲【音標】，其【拼音字母】有 o 長音字母，且其右上方帶有圓點調符（\u0358）： o͘
        letter, tone, nasal = match.groups()
        # 將 o 長音字母，轉換成【拼音字母】 oo，再附回聲調
        # replaced = f"{letter}{letter}{tone}"
        replaced = f"{letter}{tone}{letter}"
        # 重組字串
        decomposed = decomposed.replace(match.group(), replaced)
    # 依 Unicode 組合標準（NFC）重構【拼音字母與調符】，組成轉換後之【音標】
    return unicodedata.normalize('NFC', decomposed)

def tng_un_bu(im_piau: str) -> str:
    # 轉換【鼻音韻母】
    im_piau = im_piau.replace("ⁿ", "nn", 1)
    # 帶調符之白話字韻母 o͘ ，轉換為【帶韻符之 oo 韻母】
    im_piau = handle_o_dot(im_piau)

    # 解構【帶調符音標】，轉成：【無調符音標】、【聲調符號】
    letters, tone = separate_tone(im_piau)

    # 以【無調符音標】，轉換【韻母】
    sorted_keys = sorted(un_bu_mapping, key=len, reverse=True)
    for key in sorted_keys:
        if key in letters:
            letters = letters.replace(key, un_bu_mapping[key])
            break

    if tone:
        letters = apply_tone(letters, tone)

    return letters

# =========================================================================
# 【帶調符拼音】轉【帶調號拼音】
# =========================================================================

def tng_im_piau(im_piau: str, po_ci: bool = True) -> str:
    """
    將【帶調符音標】（台羅拼音/台語音標）轉換成【帶調號TLPA音標】
    :param im_piau: str - 帶調符音標
    :param po_ci: bool - 是否保留【音標】之首字母大寫
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
        if letters[0] == "u":
            letters = "U" + letters[1:]
        elif letters[0] == "i":
            letters = "I" + letters[1:]
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

def split_tlpa_im_piau(im_piau: str, po_ci: bool = False):
    # 查檢傳入之【音標】是否為帶調號TLPA音標
    if im_piau[-1] in "123456789": return im_piau
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
    elif im_piau.startswith("chh"):
        im_piau = im_piau.replace("chh", "c", 1)
    elif im_piau.startswith("ch"):
        im_piau = im_piau.replace("ch", "z", 1)

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
    un_bu = tng_un_bu(un_bu)

    # 調整聲母大小寫
    if po_ci:
        siann_bu = siann_bu.capitalize() if siann_bu else ""
    else:
        siann_bu = siann_bu.lower()

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

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


#=========================================================================
# 測試程式
#=========================================================================

def ut01():
    im_piau = "Ín"
    print(f"im_piau = {im_piau}")
    im_piau_iong_tiau_ho = tng_tiau_ho(im_piau)
    print(f"im_piau_iong_tiau_ho = {im_piau_iong_tiau_ho}")

def ut02():
    im_piau = "Ín."
    print(f"im_piau = {im_piau}")
    im_piau_iong_tiau_ho = tng_tiau_ho(im_piau)
    print(f"im_piau_iong_tiau_ho = {im_piau_iong_tiau_ho}")

def ut03():
    im_piau = "Ín\u200B"
    print(f"im_piau = {im_piau} (len = {len(im_piau)})")
    im_piau_iong_tiau_ho = tng_tiau_ho(im_piau)
    print(f"im_piau_iong_tiau_ho = {im_piau_iong_tiau_ho} (len = {len(im_piau_iong_tiau_ho)}) ")

def ut04():
    im_piau = "Ín\u200B"
    print(f"im_piau = {im_piau} (len = {len(im_piau)})")
    im_piau_cleaned = clean_im_piau(im_piau)
    print(f"im_piau_cleaned = {im_piau_cleaned} (len = {len(im_piau_cleaned)}) ")
    im_piau_iong_tiau_ho = tng_tiau_ho(im_piau)
    print(f"im_piau_iong_tiau_ho = {im_piau_iong_tiau_ho} (len = {len(im_piau_iong_tiau_ho)}) ")

def ut05():
    oo = ("o\u0300\u0358", "o\u0301\u0358", "o\u0302\u0358", "hô\u0358")
    for im_piau in oo:
        im_piau = handle_o_dot(im_piau)
        print(im_piau)

def ut06():
    oo = ("o\u0300\u0358", "o\u0301\u0358", "o\u0302\u0358", "hô\u0358")
    for im_piau in oo:
        im_piau = handle_o_dot(im_piau)
        print(im_piau)

def ut07():
    ku_cleaned = 'Kue kì lâi ê ! Tiân ôan chiong û hô put kue ?'
    # 解構【音標】組成之【句子】，變成單一【帶調符音標】清單
    im_piau_list = ku_cleaned.split()

    # 轉換成【帶調號拼音】
    converted_list = []
    for im_piau in im_piau_list:
        # 排除標點符號不進行韻母轉換
        if re.match(r'[a-zA-Zâîûêôáéíóúàèìòùāēīōūǎěǐǒǔ]+$', im_piau, re.I):
            converted_im_piau = tng_un_bu(im_piau)
        else:
            converted_im_piau = im_piau

        converted_list.append(converted_im_piau)

    print(converted_list)

def ut08():
    print("\n---------------------------------------------------")
    # 歸去來兮！田園將蕪胡不歸？​
    ku = "Kue kì lâi ê! Tiân-ôan chiong û hô put kue?​"
    # 去除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
    ku_cleaned = zing_li_zuan_ku(ku)
    # 將整段句子，解構成單一【音標】清單
    im_piau_list = [im_piau for im_piau in ku_cleaned.split()]
    # 顯示解構後的音標清單
    for im_piau in im_piau_list:
        print(im_piau, end=" ")
    print("\n---------------------------------------------------")

    # 存放轉換後的音標清單
    tlpa_im_piau_list = []
    for im_piau in im_piau_list:
        tng_uann_hau = ""
        # 排除標點符號不進行韻母轉換
        if im_piau in PUNCTUATIONS:
            # 非【音標】，視同【標點符號】，直接存入
            tng_uann_hau = im_piau
        else:
            # 將【音標】轉換為【TLPA音標】
            tlpa_im_piau_tua_tiau_hu = tng_im_piau(im_piau)
            # 將【帶調符音標】轉換為【帶調號音標】
            tng_uann_hau = tng_tiau_ho(tlpa_im_piau_tua_tiau_hu)

        tlpa_im_piau_list.append(tng_uann_hau)

    for im_piau in tlpa_im_piau_list:
        print(im_piau, end=" ")
    print("\n-----------------------------------------------------------")

    # 將音標清單轉換為句子
    sentence = normalize_sentence(tlpa_im_piau_list)
    print(f"sentence = {sentence}")

def ut09():
    print("\n-----------------------------------------------------------")
    ku = "Kue kì lâi ê! Tiân-ôan chiong û hô put kue?​"
    # 去除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
    ku_cleaned = zing_li_zuan_ku(ku)
    # 將整段句子，解構成單一【音標】清單
    im_piau_list = [im_piau for im_piau in ku_cleaned.split()]
    # 顯示解構後的音標清單
    for im_piau in im_piau_list:
        print(f"im_piau = {im_piau}：", end=" ")
        # 排除標點符號不進行韻母轉換
        if im_piau in PUNCTUATIONS:
            print("標點符號，不必判斷\n")
            continue
        if kam_si_u_tiau_hu(im_piau):
            print("帶調符音標")
        else:
            print("無調符音標")
        print("")
    print("\n-----------------------------------------------------------")

def ut10():
    # 轉換音標中【韻母】部份，不含【o͘】（oo長音）的特殊處理
    letters, tone = separate_tone('Ióng')   # 無調符音標：im_piau
    if tone:
        tiau_ho = tiau_fu_mapping[tone]
        unicode_code = f"U+{ord(tone):04X}"  # 例如 tone = '́' → 'U+0301'
        print(f"{unicode_code} ==> 調號： {tiau_ho}")
        print("-----------------------------------------------------------")

def ut11():
    # 測試 apply_tone 函數
    print(apply_tone("Iong", "\u0301"))  # 勇 → Ióng ✅
    print(apply_tone("noo", "\u0301"))   # 老 → nóo
    print(apply_tone("ngoo", "\u0302"))  # 吾 → ngôo
    print(apply_tone("tong", "\u0304"))  # 同 → tōng
    print(apply_tone("Phuan", "\u0302")) # 盤 → Puân
    print(apply_tone("liu", "\u0302"))   # 劉 → liû
    print(apply_tone("lau", "\u0302"))   # 流 → lâu

def ut12():
    #=================================================================
    # 存放轉換後的音標清單
    # tlpa_im_piau_list = []
    han_ji_list = ['罟', '勇', '輶']
    im_piau_list = ['ko͘', 'Ióng', 'Iû']
    # 顯示解構後的音標清單
    for im_piau in im_piau_list:
        print(im_piau, end=" ")
    print("\n-----------------------------------------------------------")
    for im_piau in im_piau_list:
        tng_uann_hau = ""
        # 排除標點符號不進行韻母轉換
        if im_piau in PUNCTUATIONS:
            # 非【音標】，視同【標點符號】，直接存入
            tng_uann_hau = im_piau
        else:
            # if kam_si_u_tiau_hu(im_piau):
            #     # 將【音標】轉換為【TLPA音標】
            #     tlpa_im_piau_tua_tiau_hu = tng_im_piau(im_piau)
            #     # 將【帶調符音標】轉換為【帶調號音標】
            #     tng_uann_hau = tng_tiau_ho(tlpa_im_piau_tua_tiau_hu)
            # else:
            #     tng_uann_hau = tng_tiau_ho(tlpa_im_piau_tua_tiau_hu)

            # 將【音標】轉換為【TLPA音標】
            tlpa_im_piau_tua_tiau_hu = tng_im_piau(im_piau)
            # 將【帶調符音標】轉換為【帶調號音標】
            tng_uann_hau = tng_tiau_ho(tlpa_im_piau_tua_tiau_hu)
            # tng_uann_hau_im_piau = split_tai_gi_im_piau(tng_uann_hau)
            # print(split_tlpa_im_piau(tng_uann_hau))
            print(f"im_piau = {tng_uann_hau} <-- {im_piau}")

# =========================================================
# 主程式
# =========================================================
if __name__ == "__main__":
    # 歸去來兮！田園將蕪胡不歸？​
    # ku = "Kue kì lâi ê! Tiân-ôan chiong û hô put kue?​"
    # ku = "Kue1 ki3 lai5 e5! Tian5 uan5 ziong1 u5 ho5 put4 kue1?"
    # ku = "Kue chhiong Chhiong chiong Chiong tshiong Tshiong tsiong Tsiong ńai ôan chiong hô​"
    # ku = "ńai Ńai"
    # # 去除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
    # ku_cleaned = zing_li_zuan_ku(ku)
    # # 將整段句子，解構成單一【音標】清單
    # im_piau_list = [im_piau for im_piau in ku_cleaned.split()]
    # # 顯示解構後的音標清單
    # for im_piau in im_piau_list:
    #     print(im_piau, end=" ")
    # print("\n-----------------------------------------------------------")
    print("-----------------------------------------------------------")
    # ut08()
    # ut09()
    ut12()
