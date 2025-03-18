import re
import unicodedata

# =========================================================
# 上標數字與普通數字的映射字典
# =========================================================
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

# =========================================================================
# 白話字（POJ）/台語羅馬字（TLPA）/ 台羅拼音（TL） 音標適用之聲調符號轉換
# =========================================================================
# 聲調符號對應調值的映射
tiau_fu_tng_tiau_ho_mapping_dict = {
    "\u0300": "3",  # 陰去 ò
    "\u0301": "2",  # 陰上 ó
    "\u0302": "5",  # 陽平 ô
    "\u0304": "7",  # 陽去 ō
    "\u0306": "9",  # 輕声 ŏ
    "\u030C": "6",  # 陽上 ǒ
    "\u030D": "8",  # 陽入 o̍h
}

# 調號字典
tiau_ho_mapping_dict = {
    "3": "\u0300",  # 陰去 ò
    "2": "\u0301",  # 陰上 ó
    "5": "\u0302",  # 陽平 ô
    "7": "\u0304",  # 陽去 ō
    "9": "\u0306",  # 輕声 ŏ
    "6": "\u030C",  # 陽上 ǒ
    "8": "\u030D",  # 陽入 o̍h
}

# =========================================================================
# 設定標點符號過濾
# =========================================================================
PUNCTUATIONS = (",", ".", "?", "!", ":", ";")

# =========================================================================
# 將使用聲調符號的 TLPA 拼音轉為改用調號數值的 TLPA 拼音
# =========================================================================

# 聲調符號對應表（帶調號母音 → 對應數字）
tiau_hu_mapping = {
    # a
    "á": ("a", "2"), "à": ("a", "3"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "ā": ("a", "8"), "a̋": ("a", "9"),
    # A
    "Á": ("A", "2"), "À": ("A", "3"), "Â": ("A", "5"), "Ǎ": ("A", "6"), "Ā": ("A", "7"), "Ā": ("A", "8"), "A̋": ("A", "9"),
    # e
    "é": ("e", "2"), "è": ("e", "3"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "ē": ("e", "8"), "e̋": ("e", "9"),
    # E
    "É": ("E", "2"), "È": ("E", "3"), "Ê": ("E", "5"), "Ě": ("E", "6"), "Ē": ("E", "7"), "Ē": ("E", "8"), "E̋": ("E", "9"),
    # i
    "í": ("i", "2"), "ì": ("i", "3"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "ī": ("i", "8"), "i̋": ("i", "9"),
    # I
    "Í": ("I", "2"), "Ì": ("I", "3"), "Î": ("I", "5"), "Ǐ": ("I", "6"), "Ī": ("I", "7"), "Ī": ("I", "8"), "I̋": ("I", "9"),
    # o
    "ó": ("o", "2"), "ò": ("o", "3"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "ō": ("o", "8"), "ő": ("o", "9"),
    # O
    "Ó": ("O", "2"), "Ò": ("O", "3"), "Ô": ("O", "5"), "Ǒ": ("O", "6"), "Ō": ("O", "7"), "Ō": ("O", "8"), "Ő": ("O", "9"),
    # u
    "ú": ("u", "2"), "ù": ("u", "3"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "ū": ("u", "8"), "ű": ("u", "9"),
    # U
    "Ú": ("U", "2"), "Ù": ("U", "3"), "Û": ("U", "5"), "Ǔ": ("U", "6"), "Ū": ("U", "7"), "Ū": ("U", "8"), "Ű": ("U", "9"),
    # m
    "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̄": ("m", "8"), "m̋": ("m", "9"),
    # M
    "Ḿ": ("M", "2"), "M̀": ("M", "3"), "M̂": ("M", "5"), "M̌": ("M", "6"), "M̄": ("M", "7"), "M̄": ("M", "8"), "M̋": ("M", "9"),
    # n
    "ń": ("n", "2"), "ǹ": ("n", "3"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̄": ("n", "8"), "n̋": ("n", "9"),
    # N
    "Ń": ("N", "2"), "Ǹ": ("N", "3"), "N̂": ("N", "5"), "Ň": ("N", "6"), "N̄": ("N", "7"), "N̄": ("N", "8"), "N̋": ("N", "9"),
}
def create_tiau_hu_mapping_horizontal():
    uan_im_ji_bu = ("a", "A", "e", "E", "i", "I", "o", "O", "u", "U", "m", "M", "n", "N")

    tiau_fu_tng_tiau_ho_mapping_dict = {
        "\u0301": "2",  # ˊ
        "\u0300": "3",  # ˋ
        "\u0302": "5",  # ˆ
        "\u030C": "6",  # ˇ
        "\u0304": "7",  # ˉ
        "\u0304 ": "8", # 陽入特殊
        "\u030B": "9",  # 雙上標
    }

    print("tiau_hu_mapping = {")

    for uan_im in uan_im_ji_bu:
        print(f"    # {uan_im}")
        row_items = []
        for tiau_fu, tiau_ho in tiau_fu_tng_tiau_ho_mapping_dict.items():
            uan_im_tiau = unicodedata.normalize("NFC", f"{uan_im}{tiau_fu}")
            uan_im_tiau = uan_im_tiau.replace(' ', '')
            row_items.append(f'"{uan_im_tiau}": ("{uan_im}", "{tiau_ho}")')
        print("    " + ", ".join(row_items) + ",")

    print("}")


def create_tiau_hu_mapping_vertical():
    # 基本元音與韻化輔音
    uan_im_ji_bu = ("a", "A", "e", "E", "i", "I", "o", "O", "u", "U", "m", "M", "n", "N")

    # 聲調符號與數字的對應表
    tiau_fu_tng_tiau_ho_mapping_dict = {
        "\u0301": "2",  # ˊ
        "\u0300": "3",  # ˋ
        "\u0302": "5",  # ˆ
        "\u030C": "6",  # ˇ
        "\u0304": "7",  # ˉ
        "\u0304 ": "8", # 特殊陽入調
        "\u030B": "9",  # 雙上標聲調
    }

    # 開始建立 dict
    print("tiau_hu_mapping = {")

    for uan_im in uan_im_ji_bu:
        for tiau_fu, tiau_ho in tiau_fu_tng_tiau_ho_mapping_dict.items():
            # 組合並轉換為單一字元（NFC）
            uan_im_tiau = unicodedata.normalize("NFC", f"{uan_im}{tiau_fu}")
            # 調整沒有成功組合的特殊情況（如陽入符號後面多空白的處理）
            uan_im_tiau = uan_im_tiau.replace(' ', '')
            # 輸出字典項
            print(f'    "{uan_im_tiau}": ("{uan_im}", "{tiau_ho}"),')

    print("}")
# =========================================================
# 韻母轉換
# =========================================================
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
    'onn': 'oonn',      # 雅俗通十五音：扛
    'uei': 'ue',        # 雅俗通十五音：檜
    'ueinn': 'uenn',    # 雅俗通十五音：檜
    'ur': 'u',          # 雅俗通十五音：艍
}

def un_bu_tng_huan(un_bu: str) -> str:
    """
    將輸入的韻母依照轉換字典進行轉換
    :param un_bu: str - 韻母輸入
    :return: str - 轉換後的韻母結果
    """

    # 韻母轉換，若不存在於字典中則返回原始韻母
    return un_bu_tng_huan_map_dict.get(un_bu, un_bu)

# =========================================================
# 解構音標 = 聲母 + 韻母 + 調號
# =========================================================

def replace_superscript_digits(input_str):
    return ''.join(superscript_digit_mapping.get(char, char) for char in input_str)

def split_tai_gi_im_piau(im_piau: str, po_ci: bool = False):
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
    un_bu = un_bu_tng_huan(un_bu)

    # 調整聲母大小寫
    if po_ci:
        siann_bu = siann_bu.capitalize() if siann_bu else ""
    else:
        siann_bu = siann_bu.lower()

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

# =========================================================================
# 用途：移除標點符號並轉換TLPA+拼音格式
# =========================================================================

# 確認音標的拼音字母中不帶：標點符號、控制字元
def clean_im_piau(im_piau: str) -> str:
    # 設定標點符號過濾
    PUNCTUATIONS = (",", ".", "?", "!", ":", ";", "\u200B")

    im_piau = ''.join(ji_bu for ji_bu in im_piau if ji_bu not in PUNCTUATIONS)  # 移除標點符號
    return im_piau

def tng_tiau_ho(im_piau: str) -> str:
    """
    將帶聲調符號的台語音標轉換為不帶聲調符號的台語音標（音標 + 調號）
    :param im_piau: str - 台語音標輸入
    :return: str - 轉換後的台語音標
    """
    # **遍歷所有可能的聲調符號**
    tone_number = "0"
    for tone_mark, (base_char, number) in tiau_hu_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除聲調符號，保留基本母音
            tone_number = number
            break

    # print(f"im_piau + tone_number = {im_piau + tone_number}")
    return im_piau + tone_number

def clean_tlpa(im_piau: str) -> str:
    # su_ji = im_piau[0]
    im_piau = clean_im_piau(im_piau)
    im_piau = tng_tiau_ho(im_piau)
    # 轉換 TLPA 音標使用之【聲母】及【韻母】
    siann_bu, un_bu, tiau = split_tai_gi_im_piau(im_piau)
    return f"{siann_bu}{un_bu}{tiau}"

def tng_poj_oo_iong_tiau_ho(im_piau: str) -> str:
    """
    將帶鼻化符號的白話字母 o 或 ô 轉換為帶調號的 oo + 調號
    :param im_piau: str - 白話字音標輸入
    :return: str - 轉換後的白話字音標
    """
    # 透過正規化，拆解聲調符號
    im_piau = unicodedata.normalize("NFD", im_piau)

    # 使用捕獲群組取出聲調符號，並替換成對應的調值
    # im_piau = re.sub(r"o[\u0300\u0301\u0302\u0304\u030D]?\u0358", "oo", im_piau)
    im_piau = re.sub(
        r"o([\u0300\u0301\u0302\u0304\u030D])?\u0358",
        lambda m: f"oo{tiau_fu_tng_tiau_ho_mapping_dict.get(m.group(1), '')}",
        im_piau
    )

    # Unicode NFC 正規化組合（重組聲調符號）
    im_piau = unicodedata.normalize("NFC", im_piau)
    return im_piau

def tng_poj_oo_iong_tiau_fu(im_piau: str) -> str:
    """
    將帶鼻化符號的白話字母 o 或 ô 轉換為帶調號的 oo + 調號
    :param im_piau: str - 白話字音標輸入
    :return: str - 轉換後的白話字音標
    """
    # Unicode NFD 正規化 (分離組合字元)
    im_piau = unicodedata.normalize("NFD", im_piau)

    # 使用捕獲群組取得聲調符號
    def convert(match):
        tone = match.group(1)
        return f"oo{tiau_fu_tng_tiau_ho_mapping_dict.get(tone, '')}"

    # 替換白話字母為oo，並附加聲調號
    # 找到帶鼻化符號(͘)的 o 或 ô，將其轉成對應的帶調符號 + o
    im_piau = re.sub(
        r"([aeiou])([\u0300\u0301\u0302\u0304\u030D])?\u0358",
        lambda m: f"{m.group(1)}{m.group(2) if m.group(2) else ''}o",
        im_piau
    )

    # 正規化回來（重組聲調符號）
    im_piau = unicodedata.normalize("NFC", im_piau)
    return im_piau

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
        im_piau = tng_poj_oo_iong_tiau_ho(im_piau)
        print(im_piau)

def ut06():
    oo = ("o\u0300\u0358", "o\u0301\u0358", "o\u0302\u0358", "hô\u0358")
    for im_piau in oo:
        im_piau = tng_poj_oo_iong_tiau_fu(im_piau)
        print(im_piau)

# =========================================================
# 主程式
# =========================================================
if __name__ == "__main__":
    # # 測試 split_tai_gi_im_piau 函式
    # im_piau = "Tsit8"
    # tng_uan_hau = split_tai_gi_im_piau(im_piau, po_ci=True)
    # print(tng_uan_hau)

    # tng_uan_hau = split_tai_gi_im_piau(im_piau)
    # print(tng_uan_hau)

    # # 測試 clean_tlpa 函式
    # chiu = "Chiu1"
    # chiu = clean_tlpa(chiu)
    # print(chiu)

    # ji = "Ín"
    # ji2 = split_tai_gi_im_piau(ji)
    # print(ji2)

    # ut04()


    # print("o\u0302\u0358")  # ô̘ (U+006F + U+0302 + U+0358)
    # print("\u006F\u0302\u0358")  # ô̘ (U+006F + U+0302 + U+0358)
    # print("o\u0358")

    # print("o\u0300\u0358")
    # print("o\u0301\u0358")
    # print("o\u0302\u0358")

    # ut05()
    # ut06()
    # create_tiau_hu_mapping_horizontal()

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

    def zuan_ku_zing_li(ku: str) -> list:
        """
        全句整理：移除多餘的控制字元、將 "-" 轉換成空白、將標點符號前後加上空白、移除多餘空白
        :param ku: str - 句子輸入
        :return: list - 斷詞結果
        """
        # 移除多餘的控制字元
        ku = re.sub(r"\s+", "\u200b", ku).strip()
        # 將 "-" 轉換成空白
        ku = ku.replace("-", " ")

        # 將標點符號前後加上空白
        ku = re.sub(f"([{''.join(re.escape(p) for p in PUNCTUATIONS)}])", r" \1 ", ku)

        # 移除多餘空白
        ku = re.sub(r"\s+", " ", ku).strip()

        return ku

    # 歸去來兮！田園將蕪胡不歸？​
    ku = "Kue kì lâi ê! Tiân-ôan chiong û hô put kue?​"
    ku_cleaned = zuan_ku_zing_li(ku)
    # im_piau_list = [clean_im_piau(im_piau) for im_piau in ku.split()]
    for im_piau in im_piau_list:
        print(im_piau, end=" ")