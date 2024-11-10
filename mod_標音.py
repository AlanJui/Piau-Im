import re
import sqlite3

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

def replace_superscript_digits(input_str):
    return ''.join(superscript_digit_mapping.get(char, char) for char in input_str)


def split_tai_lo(input_str):
    # 將上標數字替換為普通數字
    input_str = replace_superscript_digits(input_str)
    # 使用正則表達式匹配聲母、韻母和調號
    pattern = r'^([ptkhmnljw]?)([aeiouáéíóúâêîôûäëïöü]+)([0-9])?$'
    match = re.match(pattern, input_str)
    if match:
        siann_bu = match.group(1)
        un_bu = match.group(2)
        tiau_ho = match.group(3) if match.group(3) else '1'  # 默認調號為1
        return siann_bu, un_bu, tiau_ho
    else:
        return None, None, None


# ==========================================================
# 自「台語音標+」，分析出：聲母、韻母、聲調
# ----------------------------------------------------------
# 【台羅音標】到【台語音標】的轉換規則
# tai_loo_to_tai_gi_mapping = {
#     'tsh': 'c',
#     'ts': 'z'
# }
# for tai_loo, tai_gi in tai_loo_to_tai_gi_mapping.items():
#     tai_gi_im = tai_gi_im.replace(tai_loo, tai_gi)
# ==========================================================

# ----------------------------------------------------------
# 自「台語音標+」，分析出：聲母、韻母、聲調
# ----------------------------------------------------------
def split_tai_gi_im_piau(im_piau):
    # 聲母相容性轉換處理（將 tsh 轉換為 c；將 ts 轉換為 z）
    # zu_im = zu_im.replace("tsh", "c")   # 將 tsh 轉換為 c
    # zu_im = zu_im.replace("ts", "z")    # 將 ts  轉換為 z
    if im_piau.startswith("tsh") or im_piau.startswith("ch"):
        im_piau = im_piau.replace("tsh", "c", 1).replace("ch", "c", 1)  # 將 tsh, ch 轉換為 c
    elif im_piau.startswith("ts") or im_piau.startswith("c"):
        im_piau = im_piau.replace("ts", "z", 1).replace("c", "z", 1)  # 將 ts, c 轉換為 z

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")

    # 韻母為 m 或 ng 這種情況的正規表示式 (m\d 或 ng\d)
    un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

    result = []

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_bu_as_m_or_ng_pattern.match(im_piau):
        siann_bu = ""  # 沒有聲母
        un_bu = im_piau[:-1]  # 韻母是 m 或 ng
        tiau = im_piau[-1]  # 聲調是最後一個字符
    else:
        # 使用正規表示式來匹配聲母
        siann_bu_match = siann_bu_pattern.match(im_piau)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()  # 找到聲母
            un_bu = im_piau[len(siann_bu):-1]  # 韻母部分
        else:
            siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
            un_bu = im_piau[:-1]  # 韻母是剩下的部分，去掉最後的聲調

        tiau = im_piau[-1]  # 最後一個字符是聲調

    # 將上標數字替換為普通數字
    tiau = replace_superscript_digits(str(tiau))
    # tiau = 7 if int(tiau_ho) == 6 else int(tiau)

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result


def split_hong_im_hu_ho(hong_im_hu_ho):
    # 定義調符對應的字典
    Hong_Im_Tiau_Hu_Dict = {
        "ˋ": 2,
        "˪": 3,
        "ˊ": 5,
        "˫": 7,
        "\u02D9": 8,  # '˙'
    }

    # 編譯調符的正則表達式模式
    HongImTiauHu = re.compile(r"[ˋ˪ˊ˫˙]", re.I)

    # 定義表示第四聲的尾字元集合
    tone_4_endings = {'ㆴ', 'ㆵ', 'ㆻ', 'ㆷ'}

    # 定義聲母的集合
    sheng_mu_ji = {
        'ㄅ', 'ㄆ', 'ㆠ', 'ㄇ',
        'ㄉ', 'ㄊ', 'ㄋ', 'ㄌ',
        'ㄍ', 'ㄎ', 'ㆣ', 'ㄏ', 'ㄫ',
        'ㄗ', 'ㄘ', 'ㆡ', 'ㄙ',
        'ㄐ', 'ㄑ', 'ㆢ', 'ㄒ',
        'ㄓ', 'ㄔ', 'ㄕ', 'ㄖ',
        'ㄭ', 'ㄪ', 'ㄬ', 'ㄈ',
    }

    # 步驟一：檢查最後一個字元是否為調符
    if HongImTiauHu.match(hong_im_hu_ho[-1]):
        tiau_fu = hong_im_hu_ho[-1]
        tiau_hao = Hong_Im_Tiau_Hu_Dict[tiau_fu]
        # 移除調符，獲得無調符的方音符號
        wu_tiau_fu_hong_im_hu_ho = hong_im_hu_ho[:-1]
    else:
        # 最後沒有調符，判斷是第一聲還是第四聲
        if hong_im_hu_ho[-1] in tone_4_endings:
            tiau_hao = 4
        else:
            tiau_hao = 1
        wu_tiau_fu_hong_im_hu_ho = hong_im_hu_ho

    # 步驟四：提取聲母和韻母
    if wu_tiau_fu_hong_im_hu_ho and wu_tiau_fu_hong_im_hu_ho[0] in sheng_mu_ji:
        sheng_mu = wu_tiau_fu_hong_im_hu_ho[0]
        yun_mu = wu_tiau_fu_hong_im_hu_ho[1:]
    else:
        sheng_mu = ''
        yun_mu = wu_tiau_fu_hong_im_hu_ho

    return [sheng_mu, yun_mu, str(tiau_hao)]


def split_zu_im(zu_im):
    # 聲母相容性轉換處理（將 tsh 轉換為 c；將 ts 轉換為 z）
    # zu_im = zu_im.replace("tsh", "c")   # 將 tsh 轉換為 c
    # zu_im = zu_im.replace("ts", "z")    # 將 ts  轉換為 z
    if zu_im.startswith("tsh") or zu_im.startswith("ch"):
        zu_im = zu_im.replace("tsh", "c", 1).replace("ch", "c", 1)  # 將 tsh, ch 轉換為 c
    elif zu_im.startswith("ts") or zu_im.startswith("c"):
        zu_im = zu_im.replace("ts", "z", 1).replace("c", "z", 1)  # 將 ts, c 轉換為 z

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")

    # 韻母為 m 或 ng 這種情況的正規表示式 (m\d 或 ng\d)
    un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

    result = []

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_bu_as_m_or_ng_pattern.match(zu_im):
        siann_bu = ""  # 沒有聲母
        un_bu = zu_im[:-1]  # 韻母是 m 或 ng
        tiau = zu_im[-1]  # 聲調是最後一個字符
    else:
        # 使用正規表示式來匹配聲母
        siann_bu_match = siann_bu_pattern.match(zu_im)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()  # 找到聲母
            un_bu = zu_im[len(siann_bu):-1]  # 韻母部分
        else:
            siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
            un_bu = zu_im[:-1]  # 韻母是剩下的部分，去掉最後的聲調

        tiau = zu_im[-1]  # 最後一個字符是聲調

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

def choose_piau_im_method(piau_im, zu_im_huat, siann_bu, un_bu, tiau_ho):
    """選擇並執行對應的注音方法"""
    if zu_im_huat == "十五音":
        return piau_im.SNI_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "白話字":
        return piau_im.POJ_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台羅拼音":
        return piau_im.TL_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "閩拼方案":
        return piau_im.BP_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "方音符號":
        return piau_im.TPS_piau_im(siann_bu, un_bu, tiau_ho)
    elif zu_im_huat == "台語音標":
        siann = piau_im.Siann_Bu_Dict[siann_bu]["台語音標"] or ""
        un = piau_im.Un_Bu_Dict[un_bu]["台語音標"]
        return f"{siann}{un}{tiau_ho}"
    return ""


# ==========================================================
# 台語音標轉換為【漢字標音】之注音符號或羅馬字音標
# ==========================================================
def tlpa_tng_han_ji_piau_im(piau_im, piau_im_huat, tai_gi_im_piau):
    siann_bu, un_bu, tiau_ho = split_tai_gi_im_piau(tai_gi_im_piau)

    if siann_bu == "" or siann_bu == None:
        siann_bu = "Ø"

    han_ji_piau_im = choose_piau_im_method(
        piau_im,
        piau_im_huat,
        siann_bu,
        un_bu,
        tiau_ho
    )
    return han_ji_piau_im


# =========================================================
# 判斷是否為標點符號的輔助函數
# =========================================================
def is_punctuation(char):
    # 如果 char 是 None，直接返回 False
    if char is None:
        return False

    # 可以根據需要擴充此列表以判斷各種標點符號
    punctuation_marks = "，。！？；：、（）「」『』《》……"
    return char in punctuation_marks


# =========================================================
# 判斷是否為標點符號的輔助函數
# =========================================================
def is_valid_han_ji(char):
    if char is None:
        return False
    else:
        char = char.strip()

    punctuation_marks = "，。！？；：、（）「」『』《》……"
    return char not in punctuation_marks

# 方音符號轉換為【台語音標】
def hong_im_tng_tai_gi_im_piau(siann, un, tiau, cursor):
    """
    根據傳入的方音符號聲母、韻母、聲調，轉換成對應的台語音標
    :param siann: 聲母 (方音符號)
    :param un: 韻母 (方音符號)
    :param tiau: 聲調 (方音符號)
    :param cursor: 數據庫游標
    :return: 包含台語音標的字典
    """
    # 查詢聲母表，將方音符號的聲母轉換成台語音標
    cursor.execute("SELECT 台語音標 FROM 聲母對照表 WHERE 方音符號 = ?", (siann,))
    siann_result = cursor.fetchone()
    if siann_result:
        tai_gi_siann = siann_result[0]  # 取得台語音標
    else:
        tai_gi_siann = ''  # 無聲母的情況

    # 查詢韻母表，將方音符號的韻母轉換成台語音標
    cursor.execute("SELECT 台語音標 FROM 韻母對照表 WHERE 方音符號 = ?", (un,))
    un_result = cursor.fetchone()
    if un_result:
        tai_gi_un = un_result[0]  # 取得台語音標
    else:
        tai_gi_un = ''

    # 查詢聲調表，將方音符號的聲調轉換成台語音標
    # cursor.execute("SELECT 方音符號調符 FROM 聲調對照表 WHERE 台羅調號 = ?", (tiau,))
    # tiau_result = cursor.fetchone()
    # if tiau_result:
    #     tai_gi_tiau = tiau_result[0]  # 取得台語音標
    # else:
    #     tai_gi_tiau = ''
    tai_gi_tiau = tiau

    return {
        '台語音標': f"{tai_gi_siann}{tai_gi_un}{tai_gi_tiau}",
        '聲母': tai_gi_siann,
        '韻母': tai_gi_un,
        '聲調': tai_gi_tiau,
    }


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

# 台語音標轉換為方音符號
def TLPA_Tng_Zap_Goo_Im(siann_bu, un_bu, siann_tiau, cursor):
    """
    根據傳入的台語音標聲母、韻母、聲調，轉換成對應的方音符號
    :param siann_bu: 聲母 (台語音標)
    :param un_bu: 韻母 (台語音標)
    :param siann_tiau: 聲調 (台語音標中的數字)
    :param cursor: 數據庫游標
    :return: 包含方音符號的字典
    """

    # 如果聲母為 None、空字串或空集合符號(無聲母)，將其設為 '英'
    if siann_bu in [None, '', '∅']:  # 假設空集合符號用 '∅' 表示
        zu_im_siann_bu = '英'  # 無聲母的情況
    else:
        # 查詢聲母表，將台語音標的聲母轉換成方音符號
        cursor.execute("SELECT 十五音 FROM 聲母對照表 WHERE 台語音標 = ?", (siann_bu,))
        siann_bu_result = cursor.fetchone()
        if siann_bu_result:
            zu_im_siann_bu = siann_bu_result[0]  # 取得方音符號
        else:
            zu_im_siann_bu = '英'  # 無聲母的情況

    # 查詢韻母表，將台語音標的韻母轉換成方音符號
    cursor.execute("SELECT 十五音 FROM 韻母對照表 WHERE 台語音標 = ?", (un_bu,))
    un_bu_result = cursor.fetchone()
    if un_bu_result:
        zu_im_un_bu = un_bu_result[0]  # 取得方音符號
    else:
        zu_im_un_bu = ''

    # 查詢聲調表，將台語音標的聲調轉換成方音符號
    cursor.execute("SELECT 十五音聲調 FROM 聲調對照表 WHERE 台羅調號 = ?", (siann_tiau,))
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
        '漢字標音': f"{zu_im_un_bu}{zu_im_siann_tiau}{zu_im_siann_bu}",
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
            8: "\u02D9"
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
    }

    def __init__(self, han_ji_khoo):
        self.Siann_Bu_Dict = None
        self.Un_Bu_Dict = None
        self.init_piau_im_dict(han_ji_khoo)
        self.TL_pattern1 = re.compile(r"(uai|uan|uah|ueh|ee|ei|oo)", re.I)
        self.TL_pattern2 = re.compile(r"(o|e|a|u|i|n|m)", re.I)
        self.POJ_pattern1 = re.compile(r"(oai|oan|oah|oeh|ee|ei)", re.I)
        self.POJ_pattern2 = re.compile(r"(o|e|a|u|i|n|m)", re.I)
        self.HongImTiauHu = re.compile(r"ˋ|˪|ˊ|˫|\u02D9", re.I)

    def _init_siann_bu_dict(self, cursor):
        # 執行 SQL 查詢
        cursor.execute("SELECT * FROM 聲母對照表")

        # 獲取所有資料
        rows = cursor.fetchall()

        # 初始化字典
        siann_bu_dict = {}

        # 從查詢結果中提取資料並將其整理成一個字典
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

    def _init_un_bu_dict(self, cursor):
        # 執行 SQL 查詢
        cursor.execute("SELECT * FROM 韻母對照表")

        # 獲取所有資料
        rows = cursor.fetchall()

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
        if han_ji_khoo == "河洛話":
            db_name = 'Ho_Lok_Ue.db'
        else:
            db_name = 'Kong_Un.db'

        with sqlite3.connect(db_name) as conn:
            cursor = conn.cursor()
            self.Siann_Bu_Dict = self._init_siann_bu_dict(cursor)
            self.Un_Bu_Dict = self._init_un_bu_dict(cursor)

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

        if siann_bu == None or siann_bu == "Ø":
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

        if siann_bu == None or siann_bu == "Ø":
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
    def BP_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "閩拼方案"
        # 將「台羅八聲調」轉換成閩拼使用的調號
        tiau_ho_remap_for_BP = {
            1: 1,  # 陰平: 44
            2: 3,  # 上聲：53
            3: 5,  # 陰去：21
            4: 7,  # 上聲：53
            5: 2,  # 陽平：24
            7: 6,  # 陰入：3?
            8: 8,  # 陽入：4?
        }

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        if siann_bu == None or siann_bu == "Ø":
            siann = ""
        else:
            siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]

        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        piau_im = f"{siann}{un}"

        # 當聲母為「空白」，韻母為：i 或 u 時，調整聲母
        un_chars = list(un)
        if siann == "":
            if un_chars[0] == "i":
                siann = "y"
            elif un_chars[0] == "u":
                siann = "w"

        pattern = r"(a|oo|ere|iu|ui|ng|e|o|i|u|m)"
        searchObj = re.search(pattern, piau_im, re.M | re.I)

        if searchObj:
            found = searchObj.group(1)
            un_chars = list(found)
            idx = 0
            if found == "iu" or found == "ui":
                idx = 1
            elif found == "oo" or found == "ng":
                idx = 0
            elif found == "ere":
                idx = 2

            # 處理韻母加聲調符號
            guan_im = un_chars[idx]
            tiau = tiau_ho_remap_for_BP[int(tiau_ho)]  # 將「傳統八聲調」轉換成閩拼使用的調號
            un_chars[idx] = self.bp_un_bu_ga_tiau_ho(guan_im, tiau)
            un_str = "".join(un_chars)
            piau_im = piau_im.replace(found, un_str)

        return piau_im

    #================================================================
    # 方音符號注音（TPS）
    # TPS_mapping_dict = {
    #     "p": "ㆴ˙",
    #     "t": "ㆵ˙",
    #     "k": "ㆻ˙",
    #     "h": "ㆷ˙",
    # }
    #================================================================
    def TPS_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "方音符號"
        tiau_ho_remap_for_TPS = {
            1: "",
            2: "ˋ",
            3: "˪",
            4: "",
            5: "ˊ",
            7: "˫",
            8: "\u02D9",
        }
        TPS_piau_im_remap_dict = {
            "ㄗㄧ": "ㄐㄧ",
            "ㄘㄧ": "ㄑㄧ",
            "ㄙㄧ": "ㄒㄧ",
            "ㆡㄧ": "ㆢㄧ",
        }

        # 將上標數字替換為普通數字
        tiau_ho = replace_superscript_digits(str(tiau_ho))
        tiau_ho = 7 if int(tiau_ho) == 6 else int(tiau_ho)

        siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]
        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        tiau = self.TONE_MARKS[piau_im_huat][tiau_ho]
        piau_im = f"{siann}{un}{tiau}"

        pattern = r"(ㄗㄧ|ㄘㄧ|ㄙㄧ|ㆡㄧ)"
        searchObj = re.search(pattern, piau_im, re.M | re.I)
        if searchObj:
            key_value = searchObj.group(1)
            piau_im = piau_im.replace(key_value, TPS_piau_im_remap_dict[key_value])

        return piau_im

    #================================================================
    # 雅俗通十五音(SNI:Nga-Siok-Thong)
    #================================================================
    def SNI_piau_im(self, siann_bu, un_bu, tiau_ho):
        piau_im_huat = "十五音"
        tiau_ho_remap_for_sip_ngoo_im = {
            1: "一",
            2: "二",
            3: "三",
            4: "四",
            5: "五",
            7: "七",
            8: "八",
        }

        siann = self.Siann_Bu_Dict[siann_bu][piau_im_huat]
        un = self.Un_Bu_Dict[un_bu][piau_im_huat]
        # tiau = tiau_ho_remap_for_sip_ngoo_im[tiau_ho]
        tiau = self.TONE_MARKS[piau_im_huat][int(tiau_ho)]
        piau_im = f"{un}{tiau}{siann}"
        return piau_im
