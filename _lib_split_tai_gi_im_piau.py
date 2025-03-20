
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


# =========================================================
# 解構音標 = 聲母 + 韻母 + 調號
# 輸入之【音標】必需是【帶調號音標】
# =========================================================
def split_tai_gi_im_piau(im_piau: str) -> list:
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
    un_bu = tua_tiau_hu_un_bu_tng_uann(un_bu)

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

