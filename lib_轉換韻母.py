# 吾是老人欲趕路
# ku_cleaned = "Ngô͘ sī Nó͘ jîn beh kánn lō͘"
# ku_tng_huan_hau = "Ngôo sī Nóo jîn beh kánn lōo"

import re
import unicodedata

# =========================================================================
# 設定標點符號過濾
# =========================================================================
PUNCTUATIONS = (",", ".", "?", "!", ":", ";")

# =========================================================================
# 韻母轉換表
# =========================================================================
# 聲調符號對應表（帶調號母音 → 對應數字）
# tiau_hu_mapping = {
#     "a": ("a", "1"), "a̍": ("a", "8"), "á": ("a", "2"), "ǎ": ("a", "6"), "â": ("a", "5"), "ā": ("a", "7"), "à": ("a", "3"),
#     "e": ("e", "1"), "e̍": ("e", "8"), "é": ("e", "2"), "ě": ("e", "6"), "ê": ("e", "5"), "ē": ("e", "7"), "è": ("e", "3"),
#     "i": ("i", "1"), "i̍": ("i", "8"), "í": ("i", "2"), "ǐ": ("i", "6"), "î": ("i", "5"), "ī": ("i", "7"), "ì": ("i", "3"),
#     "o": ("o", "1"), "o̍": ("o", "8"), "ó": ("o", "2"), "ǒ": ("o", "6"), "ô": ("o", "5"), "ō": ("o", "7"), "ò": ("o", "3"),
#     "u": ("u", "1"), "u̍": ("u", "8"), "ú": ("u", "2"), "ǔ": ("u", "6"), "û": ("u", "5"), "ū": ("u", "7"), "ù": ("u", "3"),
#     "m": ("m", "1"), "m̍": ("m", "8"), "ḿ": ("m", "2"), "m̌": ("m", "6"), "m̂": ("m", "5"), "m̄": ("m", "7"), "m̀": ("m", "3"),
#     "n": ("n", "1"), "n̍": ("n", "8"), "ń": ("n", "2"), "ň": ("n", "6"), "n̂": ("n", "5"), "n̄": ("n", "7"), "ǹ": ("n", "3"),
#     "A": ("A", "1"), "A̍": ("A", "8"), "Á": ("A", "2"), "Ǎ": ("A", "6"), "Â": ("A", "5"), "Ā": ("A", "7"), "À": ("A", "3"),
#     "E": ("E", "1"), "E̍": ("E", "8"), "É": ("E", "2"), "Ě": ("E", "6"), "Ê": ("E", "5"), "Ē": ("E", "7"), "È": ("E", "3"),
#     "I": ("I", "1"), "I̍": ("I", "8"), "Í": ("I", "2"), "Ǐ": ("I", "6"), "Î": ("I", "5"), "Ī": ("I", "7"), "Ì": ("I", "3"),
#     "O": ("O", "1"), "O̍": ("O", "8"), "Ó": ("O", "2"), "Ǒ": ("O", "6"), "Ô": ("O", "5"), "Ō": ("O", "7"), "Ò": ("O", "3"),
#     "U": ("U", "1"), "U̍": ("U", "8"), "Ú": ("U", "2"), "Ǔ": ("U", "6"), "Û": ("U", "5"), "Ū": ("U", "7"), "Ù": ("U", "3"),
#     "M": ("M", "1"), "M̍": ("M", "8"), "Ḿ": ("M", "2"), "M̌": ("M", "6"), "M̂": ("M", "5"), "M̄": ("M", "7"), "M̀": ("M", "3"),
#     "N": ("N", "1"), "N̍": ("N", "8"), "Ń": ("N", "2"), "Ň": ("N", "6"), "N̂": ("N", "5"), "N̄": ("N", "7"), "Ǹ": ("N", "3"),
# }
tiau_hu_mapping = {
    "a": ("a", "1"), "á": ("a", "2"), "à": ("a", "3"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "a̍": ("a", "8"), "a̋": ("a", "9"),
    "A": ("A", "1"), "Á": ("A", "2"), "À": ("A", "3"), "Â": ("A", "5"), "Ǎ": ("A", "6"), "Ā": ("A", "7"), "A̍": ("A", "8"), "A̋": ("A", "9"),
    "e": ("e", "1"), "é": ("e", "2"), "è": ("e", "3"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "e̍": ("e", "8"), "e̋": ("e", "9"),
    "E": ("E", "1"), "É": ("E", "2"), "È": ("E", "3"), "Ê": ("E", "5"), "Ě": ("E", "6"), "Ē": ("E", "7"), "E̍": ("E", "8"), "E̋": ("E", "9"),
    "i": ("i", "1"), "í": ("i", "2"), "ì": ("i", "3"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "i̍": ("i", "8"), "i̋": ("i", "9"),
    "I": ("I", "1"), "Í": ("I", "2"), "Ì": ("I", "3"), "Î": ("I", "5"), "Ǐ": ("I", "6"), "Ī": ("I", "7"), "I̍": ("I", "8"), "I̋": ("I", "9"),
    "o": ("o", "1"), "ó": ("o", "2"), "ò": ("o", "3"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "o̍": ("o", "8"), "ő": ("o", "9"),
    "O": ("O", "1"), "Ó": ("O", "2"), "Ò": ("O", "3"), "Ô": ("O", "5"), "Ǒ": ("O", "6"), "Ō": ("O", "7"), "O̍": ("O", "8"), "Ő": ("O", "9"),
    "u": ("u", "1"), "ú": ("u", "2"), "ù": ("u", "3"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "u̍": ("u", "8"), "ű": ("u", "9"),
    "U": ("U", "1"), "Ú": ("U", "2"), "Ù": ("U", "3"), "Û": ("U", "5"), "Ǔ": ("U", "6"), "Ū": ("U", "7"), "U̍": ("U", "8"), "Ű": ("U", "9"),
    "m": ("m", "1"), "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̍": ("m", "8"), "m̋": ("m", "9"),
    "M": ("M", "1"), "Ḿ": ("M", "2"), "M̀": ("M", "3"), "M̂": ("M", "5"), "M̌": ("M", "6"), "M̄": ("M", "7"), "M̍": ("M", "8"), "M̋": ("M", "9"),
    "n": ("n", "1"), "ń": ("n", "2"), "ǹ": ("n", "3"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̍": ("n", "8"), "n̋": ("n", "9"),
    "N": ("N", "1"), "Ń": ("N", "2"), "Ǹ": ("N", "3"), "N̂": ("N", "5"), "Ň": ("N", "6"), "N̄": ("N", "7"), "N̍": ("N", "8"), "N̋": ("N", "9"),
}

un_bu_tng_huan_map_dict = {
    'ee': 'e', 'er': 'e', 'erh': 'eh', 'or': 'o', 'ere': 'ue', 'ereh': 'ueh',
    'ir': 'i', 'eng': 'ing', 'oa': 'ua', 'oe': 'ue', 'oai': 'uai', 'ei': 'e',
    'ou': 'oo', 'onn': 'oonn', 'uei': 'ue', 'ueinn': 'uenn', 'ur': 'u',
}

# 處理 o͘ 韻母特殊情況的函數
def handle_o_dot(im_piau):
    decomposed = unicodedata.normalize('NFD', im_piau)
    # 找出 o + 聲調 + 鼻化符號的特殊組合
    match = re.search(r'(o)([\u0300\u0301\u0302\u0304\u0306\u030B\u030C\u030D]?)(\u0358)', decomposed, re.I)
    if match:
        letter, tone, nasal = match.groups()
        # 轉為 oo，再附回聲調
        replaced = f"{letter}{letter}{tone}"
        # 重組字串
        decomposed = decomposed.replace(match.group(), replaced)
    return unicodedata.normalize('NFC', decomposed)

def separate_tone(s):
    """拆解帶調字母為無調字母與調號"""
    decomposed = unicodedata.normalize('NFD', s)
    letters = ''.join(c for c in decomposed if unicodedata.category(c) != 'Mn')
    tones = ''.join(c for c in decomposed if unicodedata.category(c) == 'Mn' and c != '\u0358')
    return letters, tones

def apply_tone(s, tone):
    """聲調符號重新加回第一個母音字母上"""
    vowels = 'aeiouAEIOU'
    for i, c in enumerate(s):
        if c in vowels:
            return unicodedata.normalize('NFC', s[:i+1] + tone + s[i+1:])
    return unicodedata.normalize('NFC', s[0] + tone + s[1:])

def un_bu_tng_huan(im_piau: str) -> str:
    # 處理特殊鼻化韻母 o͘
    im_piau = handle_o_dot(im_piau)

    letters, tone = separate_tone(im_piau)
    sorted_keys = sorted(un_bu_tng_huan_map_dict, key=len, reverse=True)

    for key in sorted_keys:
        if key in letters:
            letters = letters.replace(key, un_bu_tng_huan_map_dict[key])
            break

    if tone:
        letters = apply_tone(letters, tone)

    return letters

def tng_tiau_ho(im_piau: str, kan_hua: bool = False) -> str:
    """
    將帶聲調符號的台語音標轉換為不帶聲調符號的台語音標（音標 + 調號）
    :param im_piau: str - 台語音標輸入
    :param kan_hua: bool - 簡化：若是【簡化】，聲調值為 1 或 4 ，去除調號值
    :return: str - 轉換後的台語音標
    """
    # 遇標點符號，不做轉換處理，直接回傳
    if im_piau[-1] in PUNCTUATIONS:
        return im_piau

    # 將傳入【音標】字串，以標準化之 NFC 組合格式，調整【帶調符拼音字母】；
    # 令以下之處理作業，不會發生【看似相同】的【帶調符拼音字母】，其實使用
    # 不同之 Unicode 編碼
    im_piau = unicodedata.normalize("NFC", im_piau)

    # 以【元音及韻化輔音清單】，比對傳入之【音標】，找出對應之【基本拼音字母】與【調號】
    tone_number = ""
    for tone_mark, (base_char, number) in tiau_hu_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除聲調符號，保留基本母音
            break

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

def ut01():
    ku_cleaned = "Ngô͘ sī Nó͘ jîn beh kánn lō͘"
    im_piau_list = ku_cleaned.split()

    converted_list = []
    for im_piau in im_piau_list:
        if re.match(r'[a-zA-Zâîûêôáéíóúàèìòùāēīōūǎěǐǒǔ͘]+$', im_piau, re.I):
            converted_im_piau = un_bu_tng_huan(im_piau)
        else:
            converted_im_piau = im_piau

        converted_list.append(converted_im_piau)

    ku_tng_huan_hau = ' '.join(converted_list)
    print(ku_tng_huan_hau)

#===========================================================================
# 拼音帶調符的句子，轉換成無調符、使用數值標調號的句子
# 拼音帶調符的句子
# ku_tai_tiau_hu = "Kue kì lâi ê ! Tiân ôan chiong û hô put kue ?"
# 【期待輸出】：
# ku_cleaned =     "Kue1 ki3 lai5 e5 ! Tian5 uan5 ziong1 u5 ho5 put4 kue1 ?"
#===========================================================================
def ut08():
    # 拼音帶調符的句子
    ku_tai_tiau_hu = 'Kue kì lâi ê ! Tiân ôan chiong û hô put kue ?'
    print(f"音標帶調符：{ku_tai_tiau_hu}")
    # 轉換成無調符、使用數值標調號的句子
    im_piau_list = ku_tai_tiau_hu.split()

    converted_list = []
    for im_piau in im_piau_list:
        if re.match(r'[a-zA-Zâîûêôáéíóúàèìòùāēīōūǎěǐǒǔ]+$', im_piau, re.I):
            converted_im_piau = un_bu_tng_huan(im_piau)
            # 去除聲調符號，轉成數值調號
            # converted_im_piau = tng_tiau_ho(converted_im_piau, kan_hua=True)
            converted_im_piau = tng_tiau_ho(converted_im_piau)
        else:
            converted_im_piau = im_piau

        converted_list.append(converted_im_piau)

    print(f"音標用調號：{converted_list}")
    # print(converted_list)

    print("轉換成帶調號音標：")
    for im_piau in converted_list:
        print(im_piau, end=" ")
    print()

    # 正規化成標準句子格式
    sentence = normalize_sentence(converted_list)
    print(f"轉成標準句子：{sentence}")

if __name__ == '__main__':
    # ut01()
    #===========================================================================
    ut08()

    # converted_list = ['Kue', 'ki3', 'lai5', 'e5', '!', 'Tian5', 'uan5', 'chiong', 'u5', 'ho5', 'put', 'kue', '?']
    # sentence = normalize_sentence(converted_list)
    # print(sentence)
