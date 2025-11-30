"""
將帶【調符】（聲調符號）的【台羅音標】，轉換成：【無調符台羅音標】+【調號】

範例：
    tīng  => ting7
    tshù  => tshu3
    beh   => beh4
"""

import unicodedata

# 調符對映表（帶調符之元音/韻化輔音 → 不帶調符之拼音字母、調號數值）
TIAU_HU_MAPPING = {
    # a 系列
    "á": ("a", "2"), "à": ("a", "3"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "a̍": ("a", "8"), "a̋": ("a", "9"),
    "Á": ("A", "2"), "À": ("A", "3"), "Â": ("A", "5"), "Ǎ": ("A", "6"), "Ā": ("A", "7"), "A̍": ("A", "8"), "A̋": ("A", "9"),

    # e 系列
    "é": ("e", "2"), "è": ("e", "3"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "e̍": ("e", "8"), "e̋": ("e", "9"),
    "É": ("E", "2"), "È": ("E", "3"), "Ê": ("E", "5"), "Ě": ("E", "6"), "Ē": ("E", "7"), "E̍": ("E", "8"), "E̋": ("E", "9"),

    # i 系列
    "í": ("i", "2"), "ì": ("i", "3"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "i̍": ("i", "8"), "i̋": ("i", "9"),
    "Í": ("I", "2"), "Ì": ("I", "3"), "Î": ("I", "5"), "Ǐ": ("I", "6"), "Ī": ("I", "7"), "I̍": ("I", "8"), "I̋": ("I", "9"),

    # o 系列
    "ó": ("o", "2"), "ò": ("o", "3"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "o̍": ("o", "8"), "ő": ("o", "9"),
    "Ó": ("O", "2"), "Ò": ("O", "3"), "Ô": ("O", "5"), "Ǒ": ("O", "6"), "Ō": ("O", "7"), "O̍": ("O", "8"), "Ő": ("O", "9"),

    # u 系列
    "ú": ("u", "2"), "ù": ("u", "3"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "u̍": ("u", "8"), "ű": ("u", "9"),
    "Ú": ("U", "2"), "Ù": ("U", "3"), "Û": ("U", "5"), "Ǔ": ("U", "6"), "Ū": ("U", "7"), "U̍": ("U", "8"), "Ű": ("U", "9"),

    # m 系列（韻化輔音）
    "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̍": ("m", "8"), "m̋": ("m", "9"),
    "Ḿ": ("M", "2"), "M̀": ("M", "3"), "M̂": ("M", "5"), "M̌": ("M", "6"), "M̄": ("M", "7"), "M̍": ("M", "8"), "M̋": ("M", "9"),

    # n 系列（韻化輔音）
    "ń": ("n", "2"), "ǹ": ("n", "3"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̍": ("n", "8"), "n̋": ("n", "9"),
    "Ń": ("N", "2"), "Ǹ": ("N", "3"), "N̂": ("N", "5"), "Ň": ("N", "6"), "N̄": ("N", "7"), "N̍": ("N", "8"), "N̋": ("N", "9"),
}


def tiau_hu_tng_tiau_ho(im_piau: str, kan_hua: bool = False) -> list:
    """
    將帶【調符】（聲調符號）的【台羅音標】，轉換成：【無調符台羅音標】、【調號】

    Args:
        im_piau: 帶調符的台羅音標（例如：tīng, tshù, beh, i）
        kan_hua: 是否簡化（若為 True，調號 1 和 4 回傳 None）

    Returns:
        [無調符音標, 調號]
        - 預設模式：["ting", "7"], ["tshu", "3"], ["beh", "4"], ["i", "1"]
        - 簡化模式（kan_hua=True）：調號 1 和 4 回傳 None
          例如：["beh", None], ["i", None]

    範例:
        >>> tiau_hu_tng_tiau_ho("tīng")
        ["ting", "7"]
        >>> tiau_hu_tng_tiau_ho("beh")
        ["beh", "4"]
        >>> tiau_hu_tng_tiau_ho("beh", True)
        ["beh", None]
        >>> tiau_hu_tng_tiau_ho("i")
        ["i", "1"]
        >>> tiau_hu_tng_tiau_ho("i", True)
        ["i", None]
    """
    # 空字串檢查
    if not im_piau:
        return ["", None if kan_hua else ""]

    # 若音標末端已經是數字（調號），表示已轉換過，分離並回傳
    if im_piau[-1] in "123456789":
        tiau_ho = im_piau[-1]
        bo_tiau_im_piau = im_piau[:-1]
        # 簡化模式下，調號 1 和 4 回傳 None
        if kan_hua and tiau_ho in ["1", "4"]:
            return [bo_tiau_im_piau, None]
        return [bo_tiau_im_piau, tiau_ho]

    # 將傳入【音標】字串，以標準化之 NFC 組合格式，調整【帶調符拼音字母】
    # 確保【帶調符拼音字母】使用統一的 Unicode 編碼
    im_piau = unicodedata.normalize("NFC", im_piau)

    # 尋找帶調符的元音或韻化輔音，並進行轉換
    tiau_ho = "1"  # 預設調號為 1（陰平）
    bo_tiau_im_piau = im_piau  # 預設為原音標

    for tiau_mark, (base_char, number) in TIAU_HU_MAPPING.items():
        if tiau_mark in im_piau:
            bo_tiau_im_piau = im_piau.replace(tiau_mark, base_char)  # 移除調符，保留基本字母
            tiau_ho = number  # 記錄調號
            break
    else:
        # 沒有找到調符，判斷是否為入聲（末字為 h, p, t, k）
        if im_piau and im_piau[-1] in "hptk":
            tiau_ho = "4"  # 陰入調
        else:
            tiau_ho = "1"  # 陰平調

    # 依是否要【簡化】之設定，處理【調號】1 或 4 是否要略去
    if kan_hua and tiau_ho in ["1", "4"]:
        return [bo_tiau_im_piau, None]  # 簡化模式下，調號 1 和 4 回傳 None
    else:
        return [bo_tiau_im_piau, tiau_ho]


def main():
    """測試函數"""
    test_cases = {
        "聽": "tīng",
        "厝": "tshù",
        "欲": "beh",
        "媠": "súi",
        "有": "ū",
        "毋": "m̄",
        "伊": "i",
        "園": "ôan",
        "為": "ûi",
        "獨": "to̍k",
        "悟": "ngōo",
        "途": "tôo",
        "遠": "oán",
        "昨": "cha̍k",
    }

    print("=" * 80)
    print("將帶【調符】的【台羅音標】轉換成【無調符台羅音標】、【調號】")
    print("=" * 80)
    print(f"{'漢字':^6} {'帶調符音標':^15} {'無調符音標':^15} {'調號':^6}")
    print("-" * 80)

    for han_ji, im_piau in test_cases.items():
        bo_tiau_im_piau, tiau_ho = tiau_hu_tng_tiau_ho(im_piau)
        print(f"{han_ji:^6} {im_piau:^15} {bo_tiau_im_piau:^15} {tiau_ho:^6}")

    print("=" * 80)
    print("\n簡化模式測試（調號 1 和 4 回傳 None）:")
    print("=" * 80)
    print(f"{'漢字':^6} {'帶調符音標':^15} {'無調符音標':^15} {'調號':^6}")
    print("-" * 80)

    for han_ji, im_piau in test_cases.items():
        bo_tiau_im_piau, tiau_ho = tiau_hu_tng_tiau_ho(im_piau, kan_hua=True)
        tiau_ho_display = str(tiau_ho) if tiau_ho is not None else "None"
        print(f"{han_ji:^6} {im_piau:^15} {bo_tiau_im_piau:^15} {tiau_ho_display:^6}")


if __name__ == "__main__":
    main()
