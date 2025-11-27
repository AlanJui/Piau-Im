"""
將台語音標（TLPA）韻母轉換為方音符號（台語用注音符號）
"""


# 台語音標單獨韻母對映表
TLPA_UN_BU_MAP = {
    # 單元音
    "a": "ㄚ",
    "i": "ㄧ",
    "ir": "ㆨ",
    "u": "ㄨ",
    "e": "ㆤ",
    "oo": "ㆦ",
    "o": "ㄜ",
    # 鼻音韻尾（單獨）
    "m": "ㆬ",
    "n": "ㄣ",
    "ng": "ㆭ",
    # 複合韻母
    "ai": "ㄞ",
    "au": "ㄠ",
    "am": "ㆰ",
    "an": "ㄢ",
    "ang": "ㄤ",
    "om": "ㆱ",
    "ong": "ㆲ",
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


def convert_tlpa_un_bu_to_zu_im(un_bu):
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
            base_bopomo = convert_tlpa_un_bu_to_zu_im(base)
            if base_bopomo:
                return base_bopomo + bopomo_jip

    # 無法轉換則返回原字串
    return un_bu


def convert_tlpa_to_zu_im(tlpa_text, include_tone=True):
    """
    將完整台語音標轉換為方音符號

    Args:
        tlpa_text: 台語音標文字（如 "ang1", "iau5" 等）
        include_tone: 是否包含聲調符號

    Returns:
        方音符號文字
    """
    if not tlpa_text:
        return ""

    # 分離聲調（假設聲調數字在最後）
    tone = ""
    un_bu_part = tlpa_text

    if tlpa_text[-1].isdigit():
        tone = tlpa_text[-1]
        un_bu_part = tlpa_text[:-1]

    # 轉換韻母
    zu_im = convert_tlpa_un_bu_to_zu_im(un_bu_part)

    # 加上聲調符號
    if include_tone and tone in TLPA_TIAU_HU_MAP:
        zu_im += TLPA_TIAU_HU_MAP[tone]

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
        result = convert_tlpa_to_zu_im(tlpa)
        status = "✓" if result == expected else "✗"
        print(f"{tlpa:15} {result:10} {expected:10} {status:5}")

    print("\n韻母對映測試:")
    print("-" * 60)
    un_bu_tests = ["a", "ai", "ang", "inn", "iau", "uai", "iong"]
    for un_bu in un_bu_tests:
        bopomo = convert_tlpa_un_bu_to_zu_im(un_bu)
        print(f"{un_bu:10} → {bopomo}")


if __name__ == "__main__":
    main()

