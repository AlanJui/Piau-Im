import re
import sys
import unicodedata

# =========================================================================
# 將使用聲調符號的 TLPA 拼音轉為改用調號數值的 TLPA 拼音
# =========================================================================
# TLPA 聲調符號對應數值
# fmt: off
TONE_MAP = {
    "á": "2", "à": "3", "â": "5", "ǎ": "6", "ā": "7",  # a
    "é": "2", "è": "3", "ê": "5", "ě": "6", "ē": "7",  # e
    "í": "2", "ì": "3", "î": "5", "ǐ": "6", "ī": "7",  # i
    "ó": "2", "ò": "3", "ô": "5", "ǒ": "6", "ō": "7",  # o
    "ú": "2", "ù": "3", "û": "5", "ǔ": "6", "ū": "7",  # u
    "ń": "2", "ň": "6", "ñ": "5"  # 特殊鼻音
}
# fmt: on

# 用途：將 TLPA 拼音中的聲調符號轉換為數字
def convert_tlpa_tone(tlpa_word):
    tone = "1"  # 預設為陰平調
    decomposed = unicodedata.normalize("NFD", tlpa_word)  # 拆解組合字元，確保聲調分離
    base_chars = []  # 存儲純字母
    last_tone = None  # 存儲最後一個聲調符號
    has_tone_8 = False  # 用來判斷是否有聲調 8（U+030D）

    for char in decomposed:
        if char in TONE_MAP:
            last_tone = TONE_MAP[char]  # 記錄最後找到的聲調
        elif char == "\u030D":  # 檢查是否為「聲調 8」（U+030D）
            has_tone_8 = True
        elif unicodedata.category(char) != "Mn":  # 不是變音符號才存入
            base_chars.append(char)

    # 如果包含聲調 8，則強制調值為 8
    if has_tone_8:
        tone = "8"
    elif last_tone:
        tone = last_tone  # 使用最後找到的聲調

    # 若尾碼為 h/p/t/k，則屬於陰入調（4調），但 **若已確定為 8，則不改變**
    if not has_tone_8 and base_chars and base_chars[-1] in "hptk":
        tone = "4"

    return "".join(base_chars) + tone

def tiau_hu_tng_tiau_ho(tlpa_word):
    t1 = unicodedata.normalize("NFD", tlpa_word)  # 拆解聲調符號
    bo_tiau_fu = ''.join([c for c in t1 if not unicodedata.combining(c)])  # 去除變音符號

# 聲調符號對應表（帶調號母音 → 對應數字）
# fmt: off
tiau_hu_mapping = {
    "a̍": ("a", "8"), "á": ("a", "2"), "ǎ": ("a", "6"), "â": ("a", "5"), "ā": ("a", "7"), "à": ("a", "3"),
    "e̍": ("e", "8"), "é": ("e", "2"), "ě": ("e", "6"), "ê": ("e", "5"), "ē": ("e", "7"), "è": ("e", "3"),
    "i̍": ("i", "8"), "í": ("i", "2"), "ǐ": ("i", "6"), "î": ("i", "5"), "ī": ("i", "7"), "ì": ("i", "3"),
    "o̍": ("o", "8"), "ó": ("o", "2"), "ǒ": ("o", "6"), "ô": ("o", "5"), "ō": ("o", "7"), "ò": ("o", "3"),
    "u̍": ("u", "8"), "ú": ("u", "2"), "ǔ": ("u", "6"), "û": ("u", "5"), "ū": ("u", "7"), "ù": ("u", "3"),
    "m̍": ("m", "8"), "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̄": ("m", "7"),
    "n̍": ("n", "8"), "ń": ("n", "2"), "ň": ("n", "6"), "n̂": ("n", "5"), "n̄": ("n", "7")
}
# fmt: on

def tiau_hu_tng_tiau_ho(im_piau: str) -> str:
    """
    將帶有聲調符號的台羅拼音轉換為改良式【台語音標】（TLPA+）。
    """
    # **重要**：先將字串標準化為 NFC 格式，統一處理 Unicode 差異
    im_piau = unicodedata.normalize("NFC", im_piau)

    tone_number = ""

    # 1. 先處理聲調轉換
    for tone_mark, (base_char, number) in tiau_hu_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除調號，還原原始母音
            tone_number = number  # 記錄對應的聲調數字
            break  # 只會有一個聲調符號，找到就停止

    # 2. 若有聲調數字，則加到末尾
    if tone_number:
        return im_piau + tone_number

    return im_piau  # 若無聲調符號則不變更


# =========================================================================
# 設定標點符號過濾
# =========================================================================
PUNCTUATIONS = (",", ".", "?", "!", ":", ";")

def clean_tlpa(word):
    word = ''.join(ch for ch in word if ch not in PUNCTUATIONS)  # 移除標點符號
    word = unicodedata.normalize("NFD", word)  # 先正規化，拆解聲調符號
    # word = word.replace("oa", "ua")  # TLPA+ 調整，將 "oa" 變為 "ua"
    word = re.sub(r"o[\u0300\u0301\u0302\u0304\u030D]?a", "ua", word)  # 替換 "oe" 為 "ue"
    word = re.sub(r"o[\u0300\u0301\u0302\u0304\u030D]?e", "ue", word)  # 替換 "oe" 為 "ue"
    word = re.sub(r"e[\u0300\u0301\u0302\u0304\u030D]?ng", "ing", word)  # 替換 "eng" 為 "ing"
    word = re.sub(r"e[\u0300\u0301\u0302\u0304\u030D]?k", "ik", word)  # 替換 "ek" 為 "ik"
    # word = re.sub(r"ô͘", "ôo", word)  # 替換所有 `ô͘`，將 POJ `ô͘` 轉換為 TLPA `ôo`
    word = re.sub(r"o\u0302\u0358", "ôo", word)  # 替換分解後的 ô͘ (o + ̂ + 鼻音符號)

    if word.startswith("chh"):
        word = "c" + word[3:]
    elif word.startswith("ch"):
        word = "z" + word[2:]

    return unicodedata.normalize("NFC", word)  # 重新組合聲調符號

# =========================================================================


# too5 = "tô͘"
# too5 = clean_tlpa(too5)
# print(f"too5: {too5}")

# tlpa_word = "tsháu"
# tlpa_bo_tiau_hu = tiau_hu_tng_tiau_ho(tlpa_word)
# print(f"tlpa_bo_tiau_hu: {tlpa_bo_tiau_hu}")

# tlpa_char = "á"
# if tlpa_char in TONE_MAP:
#     print(f"tlpa_char: {tlpa_char} -> {TONE_MAP[tlpa_char]}")

# tlpa_word = "lo̍k"

# t1 = unicodedata.normalize("NFD", tlpa_word)  # 拆解聲調符號
# print(f"t1: {t1}")

# word_a = unicodedata.normalize("NFD", "o̍e")  # 拆解聲調符號
# print(f"word_a: {word_a}")

# word_b = unicodedata.normalize("NFC", "o̍e")  # 合併回去
# print(f"word_b: {word_b}")

# word_c = unicodedata.normalize("NFC", "lo\u030Dk")  # 拆解聲調符號
# print(f"word_c: {word_c}")

# #
# word_d = unicodedata.normalize("NFD", word_c)  # 合併回去
# print(f"word_d: {word_d}")

# t1 = unicodedata.normalize("NFD", word_c)  # 拆解聲調符號
# t2 = ''.join([c for c in t1 if not unicodedata.combining(c)])  # 去除變音符號

# =========================================================================

# x_word = "\u212B"  # Ångström符號
# print(f"x_word: {x_word}")

# x_nfd = unicodedata.normalize("NFD", x_word)  # 拆解組合字元：Å -> A（\u0041） + °（\u030A）
# print(f"x_nfd: {x_nfd} (len: {len(x_nfd)})")

# x_nfc = unicodedata.normalize("NFC", x_nfd)  # 合併回去：Å (\u00C5)
# print(f"x_nfc: {x_nfc} (len: {len(x_nfc)})")
