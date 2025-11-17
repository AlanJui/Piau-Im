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

from a003_使用漢字注音工作表製作文章純文字 import main as a003_main
from mod_file_access import save_as_new_file

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


# =========================================================
# 音標整埋工具庫
# =========================================================
# 用途：檢查是否為漢字
def is_han_ji(char):
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
    """聲調符號重新加回第一個母音字母上"""
    vowels = 'aeioumnAEIOUMN'
    for i, c in enumerate(im_piau):
        if c in vowels:
            return unicodedata.normalize('NFC', im_piau[:i+1] + tone + im_piau[i+1:])
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

def tng_im_piau(bo_tiau_hu_im_piau: str, po_ci: bool = True) -> str:
    """
    將【帶調符音標】轉換成【帶調號TLPA音標】
    :param im_piau: str - 帶調符音標
    :param po_ci: bool - 是否保留【音標】之首字母大寫
    :param kan_hua: bool - 簡化：若是【簡化】，聲調值為 1 或 4 ，去除調號值
    :return: str - 轉換後的【帶調號TLPA音標】
    """
    # 遇標點符號，不做轉換處理，直接回傳
    if bo_tiau_hu_im_piau[-1] in PUNCTUATIONS:
        return bo_tiau_hu_im_piau

    # 將傳入【音標】字串，以標準化之 NFC 組合格式，調整【帶調符拼音字母】；
    # 令以下之處理作業，不會發生【看似相同】的【帶調符拼音字母】，其實使用
    # 不同之 Unicode 編碼
    bo_tiau_hu_im_piau = unicodedata.normalize("NFC", bo_tiau_hu_im_piau)

    #---------------------------------------------------------
    # 保留【音標】之首字母
    #---------------------------------------------------------
    su_ji = bo_tiau_hu_im_piau[0]      # 保存【音標】之拼音首字母
    bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.lower()
    #---------------------------------------------------------
    # 轉換【音標】的【聲母】
    #---------------------------------------------------------
    if bo_tiau_hu_im_piau.startswith("tsh"):
        bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace("tsh", "c", 1)
    elif bo_tiau_hu_im_piau.startswith("chh"):
        bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace("chh", "c", 1)
    elif bo_tiau_hu_im_piau.startswith("ts"):
        bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace("ts", "z", 1)
    elif bo_tiau_hu_im_piau.startswith("ch"):
        bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace("ch", "z", 1)

    #---------------------------------------------------------
    # 轉換【音標】的【韻母】
    #---------------------------------------------------------
    # 轉換【鼻音韻母】
    bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace("ⁿ", "nn", 1)
    # 轉換音標中【韻母】為【o͘】（oo長音）的特殊處理
    bo_tiau_hu_im_piau = handle_o_dot(bo_tiau_hu_im_piau)

    # 轉換音標中【韻母】部份，不含【o͘】（oo長音）的特殊處理
    bo_tiau_hu_im_piau, tone = separate_tone(bo_tiau_hu_im_piau)   # 無調符音標：bo_tiau_hu_im_piau
    sorted_keys = sorted(un_bu_mapping, key=len, reverse=True)

    for key in sorted_keys:
        if key in bo_tiau_hu_im_piau:
            bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace(key, un_bu_mapping[key])
            break

    # 如若傳入之【音標】首字母為大寫，則將已轉成 "z" 或 "c" 之拼音字母改為大寫
    if su_ji.isupper():
        if bo_tiau_hu_im_piau[0] == "c":
            bo_tiau_hu_im_piau = "C" + bo_tiau_hu_im_piau[1:]
        elif bo_tiau_hu_im_piau[0] == "z":
            bo_tiau_hu_im_piau = "Z" + bo_tiau_hu_im_piau[1:]
        elif bo_tiau_hu_im_piau[0] == "u":
            bo_tiau_hu_im_piau = "U" + bo_tiau_hu_im_piau[1:]
        elif bo_tiau_hu_im_piau[0] == "i":
            bo_tiau_hu_im_piau = "I" + bo_tiau_hu_im_piau[1:]
        else:
            bo_tiau_hu_im_piau = su_ji + bo_tiau_hu_im_piau[1:]
    # 調符
    # if tone: print(f"調符：{hex(ord(tone))}")
    if tone:
        bo_tiau_hu_im_piau = apply_tone(bo_tiau_hu_im_piau, tone)

    return bo_tiau_hu_im_piau

def tng_tiau_ho(im_piau: str, kan_hua: bool = False) -> str:
    """
    將【帶調符音標】轉換為【帶調號音標】
    :param im_piau: str - 帶調符音標
    :param kan_hua: bool - 簡化：若是【簡化】，聲調值為 1 或 4 ，去除調號值
    :return: str - 帶調號音標
    """
    # 遇標點符號，不做轉換處理，直接回傳
    if im_piau[-1] in PUNCTUATIONS:
        return im_piau

    # 若【音標】末端為數值，表音標已是【帶調號拼音】，直接回傳
    u_tiau_ho = True if im_piau[-1] in "123456789" else False
    if u_tiau_ho: return im_piau

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
# 程式區域函式
# =========================================================================

# 用途：從純文字檔案讀取資料並回傳 [(漢字, TLPA), ...] 之格式
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

# =========================================================================
# 用途：將漢字及TLPA標音填入Excel指定工作表
# =========================================================================
def fill_hanzi_and_tlpa(wb, use_tiau_ho=True, filename='tmp.txt', sheet_name='漢字注音', start_row=5, piau_im_soo_zai=-2):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    text_with_tlpa = read_text_with_tlpa(filename)

    row_han_ji = start_row      # 漢字位置
    row_im_piau = row_han_ji + piau_im_soo_zai   # 標音所在: -1 ==> 自動標音； -2 ==> 人工標音
    start_col = 4   # 從D欄開始
    max_col = 18    # 最大可填入的欄位（R欄）

    col = start_col

    for han_ji_ku, im_piau_ku in text_with_tlpa:
        #------------------------------------------------------------------------------
        # 填入【漢字】
        #------------------------------------------------------------------------------
        # 由於每組漢字句子，其字數可能超過15個，以致換行，故需記錄原始位置
        org_row_han_ji = row_han_ji
        org_row_im_piau = row_im_piau
        for han_ji in han_ji_ku:
            if col > max_col:
                # 超過欄位，換到下一組行
                row_han_ji += 4
                row_im_piau += 4
                col = start_col

            sheet.cells(row_han_ji, col).value = han_ji
            sheet.cells(row_han_ji, col).select()  # 選取，畫面滾動
            col += 1  # 填入後右移一欄
            # 以下程式碼有假設：每組漢字之結尾，必有標點符號
            sheet.cells(row_han_ji, col).value = "=CHAR(10)"
        # 防漢字句子總字數有超過15個之可能，故需於此回復原始之 row_han_ji 及 row_im_piau 之 row no
        row_han_ji = org_row_han_ji
        row_im_piau = org_row_im_piau

        #------------------------------------------------------------------------------
        # 處理【音標】句子，將【字串】(String)資料轉換成【清單】(List)，使之與【漢字】一一對映
        #------------------------------------------------------------------------------
        # 整理整個句子，移除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
        im_piau_ku_cleaned = zing_li_zuan_ku(im_piau_ku)

        # 解構【音標】組成之【句子】，變成單一【帶調符音標】清單
        im_piau_list = [im_piau for im_piau in im_piau_ku_cleaned.split() if im_piau]

        # 轉換成【帶調號拼音】
        im_piau_zoo = []
        for im_piau in im_piau_list:
            # 排除標點符號不進行韻母轉換
            if im_piau in PUNCTUATIONS:
                # 若為標點符號，無需轉換
                tlpa_im_piau = im_piau
            else:
                # 符合【帶調符音標】格式者，則進行【帶調號音標】轉換
                tlpa_im_piau = tng_im_piau(im_piau)    # 完成轉換之音標 = 音標帶調號

            im_piau_zoo.append(tlpa_im_piau)

        #------------------------------------------------------------------------------
        # 填入【音標】
        #------------------------------------------------------------------------------
        col = start_col     # 重設【欄數】為： 4（D欄）
        im_piau_idx = 0
        # 執行到此，【音標】應已轉換為【帶調號之TLPA音標】
        while im_piau_idx < len(im_piau_zoo):
            if col > max_col:   # 若已填滿一行（col = 19），則需換行
                row_han_ji += 4
                row_im_piau += 4
                col = start_col
            han_ji = sheet.cells(row_han_ji, col).value
            tlpa_im_piau = im_piau_zoo[im_piau_idx]
            im_piau = ""
            if han_ji and is_han_ji(han_ji):
                # 若 cell_char 為漢字，
                if use_tiau_ho:
                    # 若設定【音標帶調號】，將 tlpa_word（音標），轉換音標格式為：【聲母】+【韻母】+【調號】
                    im_piau = tng_tiau_ho(tlpa_im_piau)
                else:
                    im_piau = tlpa_im_piau
            # 填入【音標】
            sheet.cells(row_im_piau, col).value = im_piau
            print(f"（{row_im_piau}, {col}）已填入: {han_ji} [ {im_piau} ] <-- {im_piau_zoo[im_piau_idx]}")
            im_piau_idx += 1
            col += 1

        # 更新下一組漢字及TLPA標音之位置
        row_han_ji += 4     # 漢字位置
        row_im_piau += 4    # 音標位置
        col = start_col     # 每句開始的欄位

    # 填入文章終止符號：φ
    sheet.cells(row_han_ji, start_col).value = "φ"
    logging.info(f"已將漢字及TLPA注音填入【{sheet_name}】工作表！")

# =========================================================================
# 主作業程序
# =========================================================================
def main():
    # 檢查是否有指定檔案名稱，若無則使用預設檔名
    filename = sys.argv[1] if len(sys.argv) > 1 else "_tmp.txt"
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

    fill_hanzi_and_tlpa(wb,
                        filename=filename,
                        use_tiau_ho=use_tiau_ho,
                        start_row=5,
                        piau_im_soo_zai=-2) # -1: 自動標音；-2: 人工標音

    if a003_main() != EXIT_CODE_SUCCESS:
        return EXIT_CODE_FAILURE

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
