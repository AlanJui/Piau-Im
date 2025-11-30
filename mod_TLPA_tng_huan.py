"""
台語音標（TLPA）轉換模組
"""

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
# 設定日誌
# =========================================================================
import re
import unicodedata
from typing import Optional, Tuple

from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)
from mod_TL_tiau_hu_tng_tiau_ho import tiau_hu_tng_tiau_ho
from mod_TLPA_tng_BP import convert_tlpa_to_zu_im_by_un_bu

init_logging()


# 台語音標【聲母】轉【方音符號】對映表
TLPA_ZU_IM_SIANN_MAP = {
    "ng": "ㄫ",
    "ph": "ㄆ",
    "th": "ㄊ",
    "kh": "ㄎ",
    "m": "ㄇ",
    "b": "ㆠ",
    "n": "ㄋ",
    "j": "ㆡ",
    "g": "ㆣ",
    "h": "ㄏ",
    "p": "ㄅ",
    "t": "ㄉ",
    "l": "ㄌ",
    "z": "ㄗ",
    "c": "ㄘ",
    "s": "ㄙ",
    "k": "ㄍ",
}


# 台語音標單獨韻母對映表
TLPA_ZU_IM_UN_MAP = {
    "iannh": "ㄧㆩㆷ",
    "iang": "ㄧㄤ",
    "iaunn": "ㄧㆯ",
    "iauh": "ㄧㄠㆷ",
    "iong": "ㄧㆲ",
    "iunnh": "ㄧㆫㆷ",
    "uang": "ㄨㄤ",
    "uainn": "ㄨㆮ",
    "uaih": "ㄨㄞㆷ",
    "uainnh": "ㄨㆮㆷ",
    "annh": "ㆩㆷ",
    "ang": "ㄤ",
    "ainn": "ㆮ",
    "aih": "ㄞㆷ",
    "auh": "ㄠㆷ",
    "ennh": "ㆥㆷ",
    "ing": "ㄧㄥ",
    "iann": "ㄧㆩ",
    "ia": "ㄧㄚㆷ",
    "iam": "ㄧㆰ",
    "ian": "ㄧㄢ",
    "iap": "ㄧㄚㆴ",
    "iat": "ㄧㄚㆵ",
    "iak": "ㄧㄚㆻ",
    "iau": "ㄧㄠ",
    "ionn": "ㄧㆧ",
    "ioh": "ㄧㄜㆷ",
    "iok": "ㄧㆦㆻ",
    "iunn": "ㄧㆫ",
    "ooh": "ㆦㆷ",
    "onnh": "ㆧㆷ",
    "ong": "ㆲ",
    "uann": "ㄨㆩ",
    "uah": "ㄨㄚㆷ",
    "uan": "ㄨㄢ",
    "uat": "ㄨㄚㆵ",
    "uai": "ㄨㄞ",
    "ueh": "ㄨㆤㆷ",
    "uinn": "ㄨㆪ",
    "uih": "ㄨㄧㆷ",
    "ngh": "ㆭㆷ",
    "ann": "ㆩ",
    "ah": "ㄚㆷ",
    "am": "ㆰ",
    "an": "ㄢ",
    "ap": "ㄚㆴ",
    "at": "ㄚㆵ",
    "at": "ㄚㆻ",
    "ai": "ㄞ",
    "au": "ㄠ",
    "enn": "ㆥ",
    "eh": "ㆤㆷ",
    "inn": "ㆪ",
    "ih": "ㄧㆷ",
    "im": "ㄧㆬ",
    "in": "ㄧㄣ",
    "ip": "ㄧㆴ",
    "it": "ㄧㆵ",
    "ik": "ㄧㆻ",
    "ia": "ㄧㄚ",
    "io": "ㄧㄜ",
    "iu": "ㄧㄨ",
    "oo": "ㆦ",
    "onn": "ㆧ",
    "oh": "ㄜㆷ",
    "om": "ㆱ",
    "op": "ㆦㆴ",
    "ok": "ㆦㆻ",
    "uh": "ㄨㆷ",
    "un": "ㄨㄣ",
    "ut": "ㄨㆵ",
    "ua": "ㄨㄚ",
    "ue": "ㄨㆤ",
    "ui": "ㄨㄧ",
    "mh": "ㆬㆷ",
    "ng": "ㆭ",
    "a": "ㄚ",
    "e": "ㆤ",
    "i": "ㄧ",
    "o": "ㄜ",
    "u": "ㄨ",
    "m": "ㆬ",
}

# 聲調之調名與調號對映表
TLPA_MIA_HO_MAP = {
    "陰平": "1",   # 陰平（無調號）
    "陽去": "7",   # 去
    "陰去": "3",   # 陰去
    "上聲": "2",   # 上聲
    "陽平": "5",   # 陽平
    "陰入": "4",   # 陰入（無調號）
    "陽入": "8",   # 陽入
    "輕聲": "0",   # 輕聲
}

# 聲調之調號與調名對映表
TLPA_HO_MIA_MAP = {
    "1": "陰平",   # 陰平（無調號）
    "7": "陽去",   # 去
    "3": "陰去",   # 陰去
    "2": "上聲",   # 上聲
    "5": "陽平",   # 陽平
    "4": "陰入",   # 陰入（無調號）
    "8": "陽入",   # 陽入
    "0": "輕聲",   # 輕聲
}

# 閩拚音標之【注音輸入法】：【聲調】與【調符】對照表
TLPA_ZU_IM_TIAU_HU_MAP = {
    "陰平": "",    # 陰平（無調號）
    "陽去": "˫",   # 陽去
    "陰去": "˪",   # 陰去
    "上聲": "ˋ",  # 上聲
    "陽平": "ˊ",  # 陽平
    "陰入": "",    # 陰入（無調號）
    "陽入": "˙",   # 陽入
    "輕聲": "⁰",   # 輕聲
}

# 閩拚音標之【注音輸入法】：【聲調】與【按鍵】對照表
TLPA_ZU_IM_TIAU_KIAN_MAP = {
    "陰平": ":",   # 陰平（無調號）
    "陽去": "5",   # 陽去
    "陰去": "3",   # 陰去
    "上聲": "4",   # 上聲
    "陽平": "6",   # 陽平
    "陰入": "[",   # 陰入（無調號）
    "陽入": "]",   # 陽入
    "輕聲": "7",   # 輕聲
}

#============================================================================
# 音節尾字為調號（數字）擷取函數
#============================================================================
# 常用上標轉換表（補足您可能遇到的上標字元）
_SUPERSCRIPT_MAP = {
    '\u2070': '0',  # ⁰
    '\u00B9': '1',  # ¹
    '\u00B2': '2',  # ²
    '\u00B3': '3',  # ³
    '\u2074': '4',  # ⁴
    '\u2075': '5',  # ⁵
    '\u2076': '6',  # ⁶
    '\u2077': '7',  # ⁷
    '\u2078': '8',  # ⁸
    '\u2079': '9',  # ⁹
}
_SUPER_TRANS = str.maketrans(_SUPERSCRIPT_MAP)

#============================================================================
# 是否音標尾字為調號（數字）
#============================================================================
def kam_u_tiau_ho(im_piau: str):
    """
    如果尾字是（或是上標）數字，就回傳 (im_piau_without_tiau, tiau_ho)；
    否則回傳 (normalized_im_piau, None)。

    會先把已知上標數字轉為一般數字，再檢查最後一個字元。
    """
    if not im_piau:
        return None, None

    im_piau = im_piau.strip()
    if not im_piau:
        return None, None

    # 先把上標數字轉成一般數字（若有）
    im_piau_norm = im_piau.translate(_SUPER_TRANS)

    # 若尾字為數字（單字元），擷取出來
    if im_piau_norm and im_piau_norm[-1].isdigit():
        return im_piau_norm[:-1], im_piau_norm[-1]

    return im_piau_norm, None

def split_tiau_ho(im_piau: str):
    """
    如果尾字是（或是上標）數字，就回傳 (im_piau_without_tiau, tiau_ho)；
    否則回傳 (normalized_im_piau, None)。

    會先把已知上標數字轉為一般數字，再檢查最後一個字元。
    """
    if not im_piau:
        return None, None

    # 清除前後【空白】
    im_piau = im_piau.strip()
    if not im_piau:
        return None, None

    # 先把上標數字轉成一般數字（若有）
    im_piau_norm = im_piau.translate(_SUPER_TRANS)

    # 確認傳入之【閩拚音標】符合格式=聲母+韻母+聲調=羅馬拚音字母+數字
    u_hap = re.match(r"^([a-z]+)(\d+)$", im_piau_norm)
    if not u_hap:
        # 如果不符合「全英文字母+數字」格式，就原樣回傳
        return [im_piau, None]

    # 提取：【無調音標】（聲母+韻母）和【調號】
    bo_tiau_piau_im, tiau = u_hap.group(1), u_hap.group(2)
    return bo_tiau_piau_im, tiau

# =========================================================================
# 將首字母為大寫之羅馬拼音字母轉換為小寫（只處理第一個字母）
# =========================================================================
def normalize_im_piau_case(im_piau: str) -> str:
    im_piau = unicodedata.normalize("NFC", im_piau)  # 先標準化 Unicode
    return im_piau[0].lower() + im_piau[1:] if im_piau else im_piau

#============================================================================
# 台羅音標（TL）轉換為台語音標（TLPA）
#============================================================================
def convert_tl_to_tlpa(tl_im_piau: str) -> list:
    """
    轉換台羅（TL）為台語音標（TLPA），只在單字邊界進行替換。
    """

    # 將傳入之【台語音標】離析成：【無調音標】、【調號】
    tl_im_piau, tiau_ho = split_tiau_ho(tl_im_piau)
    if tl_im_piau is None or tiau_ho is None:
        return None, None, None  # 無法處理

    #------------------------------------------------------------------------
    # 查檢【台語音標】是否符合【標準】=【聲母】+【韻母】+【調號】；若是將：【陰平】、【陰入】調，
    # 略去【調號】數值：1、4，則進行矯正
    #------------------------------------------------------------------------

    # 若輸入之【台語音標】未循【標準】，對【陰平】、【陰入】聲調，省略【調號】值：【1】/【4】
    # 則依此規則進行矯正：若【調號】（即：拼音最後一個字母）為 [ptkh]，則更正調號值為 4；
    # 若【調號】填入【韻母】之拼音字元，則將【調號】則更正為 1
    # 注意：只有在 tiau_ho 為 None 時才需要補充預設值
    if tiau_ho is None:
        tiau_mia = ""
        if tl_im_piau[-1] in ['p', 't', 'k', 'h']: # 如果【韻尾】是：入聲韻尾
            tiau_mia = '陰入'  # 聲調值為 4（陰入聲）
            tiau_ho = TLPA_MIA_HO_MAP[tiau_mia]  # 調號值設為 4
        elif tl_im_piau[-1] in ['a', 'e', 'i', 'o', 'u', 'm', 'n', 'g']:  # 如果【韻尾】是：韻母或韻化輔音
            tiau_mia = '陰平'  # 聲調值為 1（陰平聲）
            tiau_ho = TLPA_MIA_HO_MAP[tiau_mia]  # 調號值設為 1

    # 將【白話字】聲母轉換成【台語音標】（將 chh 轉換為 c；將 ch 轉換為 z）
    tl_im_piau = re.sub(r'^chh', 'c', tl_im_piau)  # `^` 表示「字串開頭」
    tl_im_piau = re.sub(r'^ch', 'z', tl_im_piau)  # `^` 表示「字串開頭」

    # 將【台羅音標】聲母轉換成【台語音標】（將 tsh 轉換為 c；將 ts 轉換為 z）
    tl_im_piau = re.sub(r'^tsh', 'c', tl_im_piau)  # `^` 表示「字串開頭」
    tl_im_piau = re.sub(r'^ts', 'z', tl_im_piau)  # `^` 表示「字串開頭」

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"^(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")

    # 韻母為 m 或 ng 這種情況的正規表示式 (m\d 或 ng\d)
    un_hua_hu_im_pattern = re.compile(r"^(m|ng)\d$")

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_hua_hu_im_pattern.match(tl_im_piau):
        siann_bu = ""  # 沒有聲母
        un_bu = tl_im_piau  # 韻母是 m 或 ng
    else:
        # 使用正規表示式來匹配聲母
        siann_bu_match = siann_bu_pattern.match(tl_im_piau)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()  # 找到聲母
            un_bu = tl_im_piau[len(siann_bu):]  # 韻母部分
        else:
            siann_bu = ""  # 沒有匹配到聲母，聲母為空字串
            un_bu = tl_im_piau  # 韻母是剩下的部分，去掉最後的聲調

    # 轉換韻母
    UN_BU_MAP = {
        "iek": "ik",
        "iao": "iau",
        "ao": "au",
        "oa": "ua",
        "oe": "ue",
        "eng": "ing",
        "ek": "ik",
    }
    un_bu = UN_BU_MAP.get(un_bu, un_bu)

    return [siann_bu, un_bu, tiau_ho]

def convert_tl_with_tiau_hu_to_tlpa(im_piau_with_tiau_hu: str) -> list:
    """
    將帶有聲調符號的台羅拼音轉換為改良式【台語音標】（TLPA+）。
    回傳：[聲母, 韻母, 調號]
    """
    # 1. 將首字母為大寫之羅馬拼音字母轉換為小寫（只處理第一個字母）
    im_piau_with_tiau_hu = normalize_im_piau_case(im_piau_with_tiau_hu)

    # 2. 將帶調符的台羅音標轉換成無調符音標+調號
    bo_tiau_im_piau, tiau_ho = tiau_hu_tng_tiau_ho(im_piau_with_tiau_hu)

    # 3. 組合成帶調號的音標，然後轉換為 TLPA
    tl_im_piau = f"{bo_tiau_im_piau}{tiau_ho}"
    siann, un, tiau = convert_tl_to_tlpa(tl_im_piau)

    return [siann, un, tiau]


#============================================================================
# 將台語音標【調號】轉換為方音符號的【調符】
#============================================================================
def convert_to_tiau_hu(tiau_ho):
    """
    將台語音標【調號】轉換為方音符號的【調符】

    Args:
        tiau_ho: 台語音標調號（如 "1", "2", "3", "5", "7", "8", "0"）

    Returns:
        方音符號調符（如 "", "ˊ", "ˇ", "ˋ", "˙", "⁰" 等）
    """

    if not tiau_ho:
        return ""

    if tiau_ho in TLPA_HO_MIA_MAP:
        # 依【調號】取得【聲調】，再依【聲調】取得【調符】
        tiau_mia = TLPA_HO_MIA_MAP[tiau_ho]
        return TLPA_ZU_IM_TIAU_HU_MAP[tiau_mia]

    return ""

#============================================================================
# 將台語音標聲母轉換為方音符號
#============================================================================
def convert_siann_bu(siann_bu):
    """
    將台語音標聲母轉換為方音符號

    Args:
        siann_bu: 台語音標聲母（如 "b", "p", "m" 等）

    Returns:
        方音符號聲母（如 "ㆠ", "ㄅ", "ㄇ" 等）
    """
    return TLPA_ZU_IM_SIANN_MAP.get(siann_bu, siann_bu)

#============================================================================
# 將台語音標【韻母】轉換為方音符號
#============================================================================
def convert_un_bu(un_bu):
    """
    將台語音標韻母轉換為方音符號

    Args:
        un_bu: 台語音標韻母（如 "ang", "iau", "inn" 等）

    Returns:
        方音符號韻母（如 "ㄤ", "ㄧㄠ", "ㆪ" 等）
    """
    if not un_bu:
        return ""

    if un_bu in TLPA_ZU_IM_UN_MAP:
        return TLPA_ZU_IM_UN_MAP[un_bu]

    # 無法轉換則返回原字串
    return un_bu

def convert_tlpa_to_zu_im_by_un_kap_tiau(un_kap_tiau, include_tiau=True):
    """
    將完整台語音標轉換為方音符號

    Args:
        tlpa_text: 台語音標文字（如 "ang1", "iau5" 等）
        include_tone: 是否包含聲調符號

    Returns:
        方音符號文字
    """
    if not un_kap_tiau:
        return ""

    # 分離聲調（假設聲調數字在最後）
    tiau = ""
    un_bu_part = un_kap_tiau

    if un_kap_tiau[-1].isdigit():
        tiau = un_kap_tiau[-1]
        un_bu_part = un_kap_tiau[:-1]

    # 轉換韻母
    zu_im = convert_tlpa_to_zu_im_by_un_bu(un_bu_part)

    # 加上聲調符號
    if include_tiau and tiau in TLPA_HO_MIA_MAP:
        zu_im += TLPA_HO_MIA_MAP[tiau]

    return zu_im

#============================================================================
# 將【閩拚音標】解構成：聲母、韻母、調號
#============================================================================
def split_tlpa_im_piau(im_piau: str):
    siann = ""
    un = ""
    tiau = ""

    # 確認傳入之【閩拚音標】不為空
    if not im_piau:
        return [siann, un, tiau]

    # 確認傳入之【閩拚音標】符合格式=聲母+韻母+聲調=羅馬拚音字母+數字
    u_hap = re.match(r"^([a-z]+)(\d+)$", im_piau)
    if not u_hap:
        # 如果不符合「全英文字母+數字」格式，就原樣回傳
        return [siann, un, tiau]

    # 提取：【無調音標】（聲母+韻母）和【調號】
    bo_tiau_im_piau, tiau = u_hap.group(1), u_hap.group(2)

    #------------------------------------------------------------------------
    # 自【無調音標】分離【聲母】與【韻母】
    #------------------------------------------------------------------------

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(ng|ph|th|kh|m|b|n|j|g|h|p|t|l|z|c|s|k)")
    # 韻化輔音(m, ng)
    un_hua_hu_im_pattern = re.compile(r"^(m|ng)\d$")
    # 【無調音標】若為韻化輔音，則聲母為空，韻母即為整段
    if un_hua_hu_im_pattern.match(bo_tiau_im_piau + tiau):
        siann = ""
        un = bo_tiau_im_piau
        return [siann, un, tiau]

    siann_match = siann_bu_pattern.match(bo_tiau_im_piau)
    if siann_match:
        # 若是比對結果，可取得【聲母】
        siann = siann_match.group(1)
        un = bo_tiau_im_piau[len(siann):]
    else:
        siann = ""
        un = bo_tiau_im_piau
    return [siann, un, tiau]

#============================================================================
# 將【閩拚音標】轉換成【注音符號】
#============================================================================
def convert_tlpa_im_piau_to_zu_im(im_piau: str):
    zu_im_siann = ""
    zu_im_un = ""
    tiau_hu = ""

    siann, un, tiau = split_tlpa_im_piau(im_piau)

    zu_im_siann = convert_siann_bu(siann)
    zu_im_un = convert_un_bu(un)
    tiau_hu = convert_to_tiau_hu(tiau)

    return [zu_im_siann, zu_im_un, tiau_hu]


#============================================================================
# 測試個案
#============================================================================
def test01():
    """測試台語音標轉方音符號"""
    test_cases = [
        ("ang1", "ㄤ"),
        ("iau2", "ㄧㄠˋ"),
        ("inn5", "ㆪˊ"),
        ("uai1", "ㄨㄞ"),
        ("iat4", "ㄧㄚㆵ"),
        ("iong1", "ㄧㆲ"),
        ("e5", "ㆤˊ"),
        ("oo1", "ㆦ"),
    ]

    print("台語音標（TLPA）轉方音符號測試:")
    print("-" * 60)
    print(f"{'TLPA':15} {'方音符號':10} {'預期':10} {'結果':5}")
    print("-" * 60)

    for tlpa, expected in test_cases:
        result = convert_tlpa_to_zu_im_by_un_kap_tiau(tlpa)
        status = "✓" if result == expected else "✗"
        print(f"{tlpa:15} {result:10} {expected:10} {status:5}")

    print("\n韻母對映測試:")
    print("-" * 60)
    un_bu_tests = ["a", "ai", "ang", "inn", "iau", "uai", "iong"]
    for un_bu in un_bu_tests:
        bopomo = convert_tlpa_to_zu_im_by_un_kap_tiau(un_bu)
        print(f"{un_bu:10} → {bopomo}")

def test02():
    test_cases = [
        ("u7", ("", "u", "7")),    # 有
        ("i1", ("", "i", "1")),    # 伊
        ("kim1", ("k", "im", "1")),
        ("ia7", ("", "ia", "7")),
        ("m7", ("", "m", "7")),     # 【毋】通
        ("hong1", ("h", "ong", "1")),
        ("kuann5", ("k", "uann", "5")),
        ("hoo7", ("h", "oo", "7")),
        ("zui2", ("z", "ui", "2")),
        ("ling2", ("l", "ing", "2")),
    ]

    print("台語音標（TLPA）分解測試:")
    print("-" * 60)
    for bp_im_piau, expected in test_cases:
        result = split_tlpa_im_piau(bp_im_piau)
        status = "✓" if result == list(expected) else "✗"
        print(f"{bp_im_piau:10} → {result} 預期: {expected} {status}")

def test03():
    test_cases = [
        ("kim1", "ㄍㄧㆬ"),
        ("ia7", "ㄧㄚ˫"),
        ("hong1", "ㄏㆲ"),
        ("kuann5", "ㄍㄨㆩˊ"),
        ("hoo7", "ㄏㆦ˫"),
        ("zui2", "ㄗㄨㄧˋ"),
        ("ling2", "ㄌㄧㄥˋ"),
    ]
    print("\n台語音標轉注音符號測試:")
    print("-" * 80)
    print(f"{'閩拚音標':15} {'注音符號':20} {'預期':20} {'結果':5}")
    print("-" * 80)
    for im_piau, expected in test_cases:
        zu_im_siann, zu_im_un, tiau_hu = convert_tlpa_im_piau_to_zu_im(im_piau)
        result = f"{zu_im_siann}{zu_im_un}{tiau_hu}"
        status = "✓" if result == expected else "✗"
        print(f"{im_piau:15} {result:20} {expected:20} {status:5}")

def test04():
    test_cases = [
        ("tēng", ["t", "ing", "7"]),
    ]
    print("\n有調符之【白話字/台羅音標】轉換測試:")
    print("-" * 80)
    print(f"{'閩拚音標':15} {'轉換結果':30} {'預期':30} {'結果':5}")
    print("-" * 80)
    for im_piau_with_tiau_hu, expected in test_cases:
        result = convert_tl_with_tiau_hu_to_tlpa(im_piau_with_tiau_hu)
        status = "✓" if result == expected else "✗"
        result_str = str(result)
        expected_str = str(expected)
        print(f"{im_piau_with_tiau_hu:15} {result_str:30} {expected_str:30} {status:5}")

def main():
    """測試台語音標轉方音符號試"""
    # test02()

    """測試【閩拚音標】轉換為【注音符號】"""
    # test03()

    """測試台語音標轉方音符號"""
    # test01()

    """測試有調符之【白話字/台羅音標】轉換"""
    test04()

if __name__ == "__main__":
    main()

