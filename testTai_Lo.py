import unicodedata

# 聲調符號對應表（帶調號母音 → 對應數字）
tone_mapping = {
    "a̍": ("a", "8"), "á": ("a", "2"), "ǎ": ("a", "6"), "â": ("a", "5"), "ā": ("a", "7"), "à": ("a", "3"),
    "e̍": ("e", "8"), "é": ("e", "2"), "ě": ("e", "6"), "ê": ("e", "5"), "ē": ("e", "7"), "è": ("e", "3"),
    "i̍": ("i", "8"), "í": ("i", "2"), "ǐ": ("i", "6"), "î": ("i", "5"), "ī": ("i", "7"), "ì": ("i", "3"),
    "o̍": ("o", "8"), "ó": ("o", "2"), "ǒ": ("o", "6"), "ô": ("o", "5"), "ō": ("o", "7"), "ò": ("o", "3"),
    "u̍": ("u", "8"), "ú": ("u", "2"), "ǔ": ("u", "6"), "û": ("u", "5"), "ū": ("u", "7"), "ù": ("u", "3"),
    "m̍": ("m", "8"), "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̄": ("m", "7"),
    "n̍": ("n", "8"), "ń": ("n", "2"), "ň": ("n", "6"), "n̂": ("n", "5"), "n̄": ("n", "7")
}

# 聲母轉換規則（台羅拼音 → 台語音標+）
initials_mapping = {
    "tsh": "c",
    "ts": "z"
}

def convert_tai_lo_to_tlpa_plus(im_piau: str) -> str:
    """
    先將台羅拼音的聲調符號轉換為 TLPA 數字聲調，然後轉換聲母為台語音標+（TLPA+）。
    """
    # **重要**：先將字串標準化為 NFC 格式，統一處理 Unicode 差異
    im_piau = unicodedata.normalize("NFC", im_piau)

    tone_number = ""

    # 1. 先處理聲調轉換
    for tone_mark, (base_char, number) in tone_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除調號，還原原始母音
            tone_number = number  # 記錄對應的聲調數字
            break  # 只會有一個聲調符號，找到就停止

    # 2. 將聲母轉換為 TLPA+
    for tai_lo, tlpa in initials_mapping.items():
        if im_piau.startswith(tai_lo):  # 只轉換開頭的聲母部分
            im_piau = tlpa + im_piau[len(tai_lo):]
            break

    # 3. 若有聲調數字，則加到末尾
    if tone_number:
        return im_piau + tone_number

    return im_piau  # 若無聲調符號則不變更

if __name__ == "__main__":
    # 測試
    test_cases = ["lio̍k", "tāi", "bô", "siâu", "lâi", "pò", "tshi̍t", "tsuan", "giâm", "ló"]
    converted = [convert_tai_lo_to_tlpa_plus(word) for word in test_cases]

    # 顯示轉換結果
    for original, converted_word in zip(test_cases, converted):
        print(f"{original} → {converted_word}")
