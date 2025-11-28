"""
將台語音標（TLPA）韻母轉換為方音符號（台語用注音符號）
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
tlpa_tng_zu_im_siann_bu_map = {
    "b":"ㆠ",
    "p":"ㄅ",
    "ph":"ㄆ",
    "m":"ㄇ",
    "n":"ㄋ",
    "t":"ㄉ",
    "th":"ㄊ",
    "l":"ㄌ",
    "g":"ㆣ",
    "k":"ㄍ",
    "kh":"ㄎ",
    "ng":"ㄫ",
    "j":"ㆡ",
    "z":"ㄗ",
    "c":"ㄘ",
    "s":"ㄙ",
    "h":"ㄏ",
}

# 台語音標單獨韻母對映表
TLPA_UN_BU_MAP = {
    # 單元音
    "oo": "ㆦ",
    "o": "ㄜ",
    "a": "ㄚ",
    "i": "ㄧ",
    "ir": "ㆨ",
    "u": "ㄨ",
    "e": "ㆤ",
    # 鼻音韻尾（單獨）
    "ng": "ㆭ",
    "n": "ㄣ",
    "m": "ㆬ",
    # 複合韻母
    "ang": "ㄤ",
    "ong": "ㆲ",
    "ai": "ㄞ",
    "au": "ㄠ",
    "am": "ㆰ",
    "an": "ㄢ",
    "om": "ㆱ",
    # 鼻化韻母
    "ann": "ㆩ",
    "inn": "ㆪ",
    "unn": "ㆫ",
    "enn": "ㆥ",
    "onn": "ㆧ",
    "ainn": "ㆮ",
    "aunn": "ㆯ",
}

# 複合韻母對映表（含介音）
TLPA_COMPOSITE_UN_BU_MAP = {
    # i 介音開頭
    "iau": "ㄧㄠ",
    "ia": "ㄧㄚ",
    "io": "ㄧㄜ",
    "iu": "ㄧㄨ",
    "iang": "ㄧㄤ",
    "ian": "ㄧㄢ",
    "iam": "ㄧㆰ",
    "ing": "ㄧㄥ",
    "in": "ㄧㄣ",
    "im": "ㄧㆬ",
    "iong": "ㄧㆲ",
    # i 鼻化韻母
    "iaunn": "ㄧㆯ",
    "iann": "ㄧㆩ",
    "ionn": "ㄧㆧ",
    "iunn": "ㄧㆫ",
    # u 介音開頭
    "uai": "ㄨㄞ",
    "ua": "ㄨㄚ",
    "ui": "ㄨㄧ",
    "ue": "ㄨㆤ",
    "uang": "ㄨㄤ",
    "uan": "ㄨㄢ",
    "un": "ㄨㄣ",
    # u 鼻化韻母
    "uainn": "ㄨㆮ",
    "uann": "ㄨㆩ",
    "uinn": "ㄨㆪ",
    "uenn": "ㄨㆥ",
}

# 入聲韻尾對映表
TLPA_JIP_SIANN_MAP = {
    "p": "ㆴ",
    "t": "ㆵ",
    "k": "ㆻ",
    "h": "ㆷ",
}

# 聲調符號對映表
TLPA_TIAU_HU_MAP = {
    "1": "",    # 陰平（無調號）
    "2": "ˊ",   # 陽平
    "3": "ˇ",   # 上聲
    "5": "ˋ",   # 陰去
    "7": "",    # 陰入（無調號）
    "8": "ˊ",   # 陽入
    "0": "⁰",   # 輕聲
}

# 閩拚音標之【注音輸入法】：【調符】與【調號】對照表
bp_zu_im_hu_tiau_map = {
    '˫': '6',   # 陽去
    '˪': '5',   # 陰去
    'ˋ': '3',  # 陰上
    'ˊ': '2',  # 陽平
    '˙': '8',   # 陽入
}

# 閩拚音標之【注音輸入法】：【調號】與【按鍵】對照表
bp_zu_im_tiau_key_map = {
    '1': ':',   # 陰平
    '6': '5',   # 陽去
    '5': '3',   # 陰去
    '3': '4',   # 陰上
    '2': '6',   # 陽平
    '7': '[',   # 陰入
    '8': ']',   # 陽入
    '0': '7',   # 輕聲 ⁰
}

# 台語音標之【注音輸入法】：【調符】與【調號】對照表
tlpa_zu_im_hu_tiau_map = {
    '˫': '6',   # 陽去
    '˪': '5',   # 陰去
    'ˋ': '3',  # 陰上
    'ˊ': '2',  # 陽平
    '˙': '8',   # 陽入
}

# 台語音標之【拚音輸入法】：【調號】與【按鍵】對照表
tlpa_zu_im_tiau_key_map = {
    '1': ':',   # 陰平
    '7': '5',   # 陽去
    '3': '3',   # 陰去
    '2': '4',   # 陰上
    '5': '6',   # 陽平
    '4': '[',   # 陰入
    '8': ']',   # 陽入
    '0': '7',   # 輕聲 ⁰
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
def convert_tlpa_to_zu_im_by_tiau(tiau_ho):
    """
    將台語音標【調號】轉換為方音符號的【調符】

    Args:
        tiau_ho: 台語音標調號（如 "1", "2", "3", "5", "7", "8", "0"）

    Returns:
        方音符號調符（如 "", "ˊ", "ˇ", "ˋ", "˙", "⁰" 等）
    """
    # 台語音標之【調號】與【聲調】對照表
    tlpa_ho_kap_tiau_map = {
        '1': '陰平',   # 陰平
        '7': '陽去',   # 陽去
        '3': '陰去',   # 陰去
        '2': '陰上',   # 陰上
        '5': '陽平',   # 陽平
        '4': '陰入',   # 陰入
        '8': '陽入',   # 陽入
        '0': '輕聲',   # 輕聲 ⁰
    }
    # 台語音標之【聲調】與【調符】對照表
    tlpa_tiau_kap_hu_map = {
        '陰平': '',    # 陰平
        '陽去': '˫',   # 陽去
        '陰去': '˪',   # 陰去
        '陰上': 'ˋ',  # 陰上
        '陽平': 'ˊ',  # 陽平
        '陰入': '',    # 陰入
        '陽入': '˙',   # 陽入
        '輕聲': '⁰',   # 輕聲 ⁰
    }

    if not tiau_ho:
        return ""

    if tiau_ho in tlpa_ho_kap_tiau_map:
        # 依【調號】取得【聲調】，再依【聲調】取得【調符】
        tiau_mia = tlpa_ho_kap_tiau_map[tiau_ho]
        return tlpa_tiau_kap_hu_map[tiau_mia]

    return ""

#============================================================================
# 將台語音標聲母轉換為方音符號
#============================================================================
def convert_tlpa_to_zu_im_by_siann_bu(siann_bu):
    """
    將台語音標聲母轉換為方音符號

    Args:
        siann_bu: 台語音標聲母（如 "b", "p", "m" 等）

    Returns:
        方音符號聲母（如 "ㆠ", "ㄅ", "ㄇ" 等）
    """
    return tlpa_tng_zu_im_siann_bu_map.get(siann_bu, siann_bu)


#============================================================================
# 將台語音標【韻母】轉換為方音符號
#============================================================================
def convert_tlpa_to_zu_im_by_un_bu(un_bu):
    """
    將台語音標韻母轉換為方音符號

    Args:
        un_bu: 台語音標韻母（如 "ang", "iau", "inn" 等）

    Returns:
        方音符號韻母（如 "ㄤ", "ㄧㄠ", "ㆪ" 等）
    """
    if not un_bu:
        return ""

    # 先檢查複合韻母（較長的優先）
    if un_bu in TLPA_COMPOSITE_UN_BU_MAP:
        return TLPA_COMPOSITE_UN_BU_MAP[un_bu]

    # 再檢查單獨韻母
    if un_bu in TLPA_UN_BU_MAP:
        return TLPA_UN_BU_MAP[un_bu]

    # 處理入聲韻尾
    # 例如：iat -> ia + t -> ㄧㄚ + ㆵ
    for jip_siann_un_bue, bopomo_jip in TLPA_JIP_SIANN_MAP.items():
        if un_bu.endswith(jip_siann_un_bue):
            base = un_bu[:-len(jip_siann_un_bue)]
            base_bopomo = convert_tlpa_to_zu_im_by_un_bu(base)
            if base_bopomo:
                return base_bopomo + bopomo_jip

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
    if include_tiau and tiau in TLPA_TIAU_HU_MAP:
        zu_im += TLPA_TIAU_HU_MAP[tiau]

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

