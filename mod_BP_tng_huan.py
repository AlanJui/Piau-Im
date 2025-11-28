"""
閩拚音標（BP）轉換模組
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
from mod_logging import (
    init_logging,
    logging_exc_error,
    logging_exception,
    logging_process_step,
    logging_warning,
)

init_logging()


# 台語音標【聲母】轉【方音符號】對映表
BP_ZU_IM_SIANN_MAP = {
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
BP_ZU_IM_UN_MAP = {
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
BP_ZU_IM_HO_TIAU_MAP = {
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
BP_ZU_IM_TIAU_HU_MAP = {
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
BP_ZU_IM_TIAU_KIAN_MAP = {
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

    if tiau_ho in BP_ZU_IM_HO_TIAU_MAP:
        # 依【調號】取得【聲調】，再依【聲調】取得【調符】
        tiau_mia = BP_ZU_IM_HO_TIAU_MAP[tiau_ho]
        return BP_ZU_IM_TIAU_HU_MAP[tiau_mia]

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
    return BP_ZU_IM_SIANN_MAP.get(siann_bu, siann_bu)


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

    if un_bu in BP_ZU_IM_UN_MAP:
        return BP_ZU_IM_UN_MAP[un_bu]

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
    if include_tiau and tiau in BP_ZU_IM_HO_TIAU_MAP:
        zu_im += BP_ZU_IM_HO_TIAU_MAP[tiau]

    return zu_im


def main():
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
        result = convert_tlpa_to_zu_im_by_un_bu(tlpa)
        status = "✓" if result == expected else "✗"
        print(f"{tlpa:15} {result:10} {expected:10} {status:5}")

    print("\n韻母對映測試:")
    print("-" * 60)
    un_bu_tests = ["a", "ai", "ang", "inn", "iau", "uai", "iong"]
    for un_bu in un_bu_tests:
        bopomo = convert_tlpa_to_zu_im_by_un_bu(un_bu)
        print(f"{un_bu:10} → {bopomo}")


if __name__ == "__main__":
    main()

