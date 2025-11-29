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

from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)
from mod_TLPA_tng_BP import convert_tlpa_to_zu_im_by_un_bu

init_logging()


# 台語音標【聲母】轉【方音符號】對映表
TLPA_ZU_IM_SIANN_MAP = {
    "bb":"ㆠ",
    "b":"ㄅ",
    "p":"ㄆ",
    "bbn":"ㄇ",
    "ln":"ㄋ",
    "d":"ㄉ",
    "t":"ㄊ",
    "l":"ㄌ",
    "gg":"ㆣ",
    "g":"ㄍ",
    "k":"ㄎ",
    "ggn":"ㄫ",
    "zz":"ㆡ",
    "z":"ㄗ",
    "c":"ㄘ",
    "s":"ㄙ",
    "h":"ㄏ",
    "y": "",  # 零聲母
    "w": "",  # 零聲母
}


# 台語音標單獨韻母對映表
TLPA_ZU_IM_UN_MAP = {
    "niah": "ㄧㆩㆷ",
    "iang": "ㄧㄤ",
    "niao": "ㄧㆯ",
    "iaoh": "ㄧㄠㆷ",
    "iong": "ㄧㆲ",
    "niuh": "ㄧㆫㆷ",
    "uang": "ㄨㄤ",
    "nuai": "ㄨㆮ",
    "uaih": "ㄨㄞㆷ",
    "uaih": "ㄨㆮㆷ",
    "nah": "ㆩㆷ",
    "ang": "ㄤ",
    "nai": "ㆮ",
    "aih": "ㄞㆷ",
    "aoh": "ㄠㆷ",
    "neh": "ㆥㆷ",
    "ing": "ㄧㄥ",
    "nia": "ㄧㆩ",
    "iah": "ㄧㄚㆷ",
    "iam": "ㄧㆰ",
    "ian": "ㄧㄢ",
    "iap": "ㄧㄚㆴ",
    "iat": "ㄧㄚㆵ",
    "iak": "ㄧㄚㆻ",
    "iao": "ㄧㄠ",
    "nio": "ㄧㆧ",
    "ioh": "ㄧㄜㆷ",
    "iok": "ㄧㆦㆻ",
    "niu": "ㄧㆫ",
    "ooh": "ㆦㆷ",
    "noh": "ㆧㆷ",
    "ong": "ㆲ",
    "nua": "ㄨㆩ",
    "uah": "ㄨㄚㆷ",
    "uan": "ㄨㄢ",
    "uat": "ㄨㄚㆵ",
    "uai": "ㄨㄞ",
    "ueh": "ㄨㆤㆷ",
    "nui": "ㄨㆪ",
    "uih": "ㄨㄧㆷ",
    "ngh": "ㆭㆷ",
    "na": "ㆩ",
    "ah": "ㄚㆷ",
    "am": "ㆰ",
    "an": "ㄢ",
    "ap": "ㄚㆴ",
    "at": "ㄚㆵ",
    "ak": "ㄚㆻ",
    "ai": "ㄞ",
    "ao": "ㄠ",
    "ne": "ㆥ",
    "eh": "ㆤㆷ",
    "ni": "ㆪ",
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
    "no": "ㆧ",
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

# 聲調符號對映表
TLPA_ZU_IM_HO_TIAU_MAP = {
    "1": "陰平",   # 陰平（無調號）
    "6": "陽去",   # 去
    "5": "陰去",   # 陰去
    "3": "上聲",   # 上聲
    "2": "陽平",   # 陽平
    "7": "陰入",   # 陰入（無調號）
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
    "陰平": ":",    # 陰平（無調號）
    "陽去": "5",   # 陽去
    "陰去": "3",   # 陰去
    "上聲": "4",  # 上聲
    "陽平": "6",  # 陽平
    "陰入": "[",    # 陰入（無調號）
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

    if tiau_ho in TLPA_ZU_IM_HO_TIAU_MAP:
        # 依【調號】取得【聲調】，再依【聲調】取得【調符】
        tiau_mia = TLPA_ZU_IM_HO_TIAU_MAP[tiau_ho]
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
    if include_tiau and tiau in TLPA_ZU_IM_HO_TIAU_MAP:
        zu_im += TLPA_ZU_IM_HO_TIAU_MAP[tiau]

    return zu_im


#============================================================================
# 將【閩拚音標】解構成：聲母、韻母、調號
#============================================================================
def split_bp_im_piau(bp_im_piau: str):
    siann = ""
    un = ""
    tiau = ""

    # 確認傳入之【閩拚音標】不為空
    if not bp_im_piau:
        return [siann, un, tiau]

    # 確認傳入之【閩拚音標】符合格式=聲母+韻母+聲調=羅馬拚音字母+數字
    u_hap = re.match(r"^([a-z]+)(\d+)$", bp_im_piau)
    if not u_hap:
        # 如果不符合「全英文字母+數字」格式，就原樣回傳
        return [siann, un, tiau]

    # 提取：【無調音標】（聲母+韻母）和【調號】
    bo_tiau_piau_im, tiau = u_hap.group(1), u_hap.group(2)

    #------------------------------------------------------------------------
    # 自【無調音標】分離【聲母】與【韻母】
    #------------------------------------------------------------------------

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(bb|b|p|m|n|d|t|l|gg|g|k|ng|zz|z|c|s|h|y|w)")
    # 韻化輔音(m, ng)
    un_hua_hu_im_pattern = re.compile(r"^(m|ng)\d$")
    # 【無調音標】若為韻化輔音，則聲母為空，韻母即為整段
    if un_hua_hu_im_pattern.match(bo_tiau_piau_im + tiau):
        siann = ""
        un = bo_tiau_piau_im
        return [siann, un, tiau]

    siann_match = siann_bu_pattern.match(bo_tiau_piau_im)
    if siann_match:
        # 若是比對結果，可取得【聲母】
        siann = siann_match.group(1)
        un = bo_tiau_piau_im[len(siann):]
    else:
        siann = ""
        un = bo_tiau_piau_im
    return [siann, un, tiau]


#============================================================================
# 將【閩拚音標】轉換成【注音符號】
#============================================================================
def convert_bp_im_piau_to_zu_im(bp_im_piau: str):
    zu_im_siann = ""
    zu_im_un = ""
    tiau_hu = ""

    siann, un, tiau = split_bp_im_piau(bp_im_piau)

    if siann == "y":
        siann = ""
        if un.startswith("i") and len(un) == 1:
            un = "i"
        else:
            un = f"y{un}"
            if un[1] in ["i", "e", "a", "o", "u"]:
                un = un.replace("y", "i", 1)
    elif siann == "w":
        siann = ""
        if un.startswith("u") and len(un) == 1:
            un = "u"
        else:
            un = f"w{un}"
            if un[1] in ["i", "e", "a", "o", "u"]:
                un = un.replace("w", "u", 1)

    zu_im_siann = convert_siann_bu(siann)
    zu_im_un = convert_un_bu(un)
    tiau_hu = convert_to_tiau_hu(tiau)
    # bp_zu_im = f"{zu_im_siann}{zu_im_un}{tiau_hu}"

    return [zu_im_siann, zu_im_un, tiau_hu]


#============================================================================
# 測試個案
#============================================================================
def test01():
    """測試台語音標轉方音符號"""
    test_cases = [
        ("ang1", "ㄤ"),
        ("iau5", "ㄧㄠˋ"),
        ("inn2", "ㆪˊ"),
        ("uai7", "ㄨㄞ"),
        ("iat4", "ㄧㄚㆵ"),
        ("iong1", "ㄧㆲ"),
        ("e3", "ㆤˇ"),
        ("oo7", "ㆦ"),
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
        ("wu7", ("w", "u", "7")),    # 有
        ("yi1", ("y", "i", "1")),    # 伊
        ("gim1", ("g", "im", "1")),
        ("ya6", ("y", "a", "6")),
        ("m7", ("", "m", "7")),     # 【毋】通
        ("hong1", ("h", "ong", "1")),
        ("gnua2", ("g", "nua", "2")),
        ("hoo6", ("h", "oo", "6")),
        ("zui3", ("z", "ui", "3")),
        ("ling3", ("l", "ing", "3")),
    ]

    print("台語音標（BP）分解測試:")
    print("-" * 60)
    for bp_im_piau, expected in test_cases:
        result = split_bp_im_piau(bp_im_piau)
        status = "✓" if result == list(expected) else "✗"
        print(f"{bp_im_piau:10} → {result} 預期: {expected} {status}")

def main():
    """測試台語音標轉方音符號"""
    # test01()

    """測試台語音標轉方音符號試"""
    test02()

    """測試【閩拚音標】轉換為【注音符號】"""
    test_cases = [
        ("gim1", "ㄍㄧㆬ"),
        ("ya6", "ㄧㄚ˫"),
        ("hong1", "ㄏㆲ"),
        ("gnua2", "ㄍㄨㆩˊ"),
        ("hoo6", "ㄏㆦ˫"),
        ("zui3", "ㄗㄨㄧˋ"),
        ("ling3", "ㄌㄧㄥˋ"),
    ]
    print("\n閩拚音標轉注音符號測試:")
    print("-" * 80)
    print(f"{'閩拚音標':15} {'注音符號':20} {'預期':20} {'結果':5}")
    print("-" * 80)
    for bp_im_piau, expected in test_cases:
        zu_im_siann, zu_im_un, tiau_hu = convert_bp_im_piau_to_zu_im(bp_im_piau)
        result = f"{zu_im_siann}{zu_im_un}{tiau_hu}"
        status = "✓" if result == expected else "✗"
        print(f"{bp_im_piau:15} {result:20} {expected:20} {status:5}")

if __name__ == "__main__":
    main()

