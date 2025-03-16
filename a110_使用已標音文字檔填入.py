# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import re
import sys
import unicodedata

import xlwings as xw

# =========================================================================
# 將使用聲調符號的 TLPA 拼音轉為改用調號數值的 TLPA 拼音
# =========================================================================

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

    tone_number = "1"

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
# 程式區域函式
# =========================================================================

# =========================================================================
# 設定標點符號過濾
# =========================================================================
PUNCTUATIONS = (",", ".", "?", "!", ":", ";")

# 用途：從純文字檔案讀取資料並回傳 [(漢字, TLPA), ...] 之格式
def read_text_with_tlpa(filename):
    text_with_tlpa = []
    with open(filename, "r", encoding="utf-8") as f:
        # 先移除 `\u200b`，確保不會影響 TLPA 拼音對應
        lines = [re.sub(r"[\u200b]", "", line.strip()) for line in f if line.strip() and not line.startswith("zh.wikipedia.org")]

    for i in range(0, len(lines), 2):
        hanzi = lines[i]
        tlpa = lines[i + 1].replace("-", " ")  # 替換 "-" 為空白字元
        text_with_tlpa.append((hanzi, tlpa))

    return text_with_tlpa


# 用途：檢查是否為漢字
def is_hanzi(char):
    return 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(char, '')

# =========================================================================
# 用途：移除標點符號並轉換TLPA+拼音格式
# =========================================================================
def clean_tlpa(word):
    word = ''.join(ch for ch in word if ch not in PUNCTUATIONS)  # 移除標點符號
    word = unicodedata.normalize("NFD", word)  # 先正規化，拆解聲調符號

     # **新增鼻音處理：將 `ⁿ`（U+207F）轉換為 `nn`**
    word = word.replace("ⁿ", "nn")

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
# 用途：將漢字及TLPA標音填入Excel指定工作表
# =========================================================================
def fill_hanzi_and_tlpa(wb, use_tiau_ho=True, filename='tmp.txt', sheet_name='漢字注音', start_row=5, piau_im_row=-2):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    text_with_tlpa = read_text_with_tlpa(filename)

    row_hanzi = start_row      # 漢字位置
    row_tlpa = row_hanzi + piau_im_row   # TLPA位置: -1 ==> 自動標音； -2 ==> 人工標音
    for idx, (hanzi, tlpa) in enumerate(text_with_tlpa):
        # 漢字逐字填入（從D欄開始）
        for col_idx, char in enumerate(hanzi):
            col = 4 + col_idx  # D欄是第4欄
            sheet.cells(row_hanzi, col).value = char
            sheet.cells(row_hanzi, col).select()  # 每字填入後選取以便畫面滾動

        # TLPA逐詞填入（從D欄開始），檢查下方儲存格是否為漢字
        tlpa_words = [clean_tlpa(word) for word in tlpa.split()]
        col = 4
        word_idx = 0

        while word_idx < len(tlpa_words):
            cell_char = sheet.cells(row_hanzi, col).value
            if cell_char and is_hanzi(cell_char):
                tlpa_word = tlpa_words[word_idx]
                if tlpa_word in PUNCTUATIONS:
                    # 若讀入之TLPA音標為標點符號，則音標儲存入空字串
                    tlpa_word = ""
                else:
                    # 若讀入之TLPA音標非標點符號，且使用標音格式二，則轉換為【聲母】+【韻母】+【調號】
                    if use_tiau_ho: tlpa_word = tiau_hu_tng_tiau_ho(tlpa_word)
                sheet.cells(row_tlpa, col).value = tlpa_word
                word_idx += 1
                print(f"（{row_tlpa}, {col}）已填入: {cell_char} - {tlpa_words[word_idx-1]}")
            col += 1

        # 完成一組漢字及TLPA標音後，需在儲存格存入換行符號
        if word_idx == len(tlpa_words):
            if col >= 18:   # 若已填滿一行（col = 19），則需換行
                col = 4
                row_hanzi += 4

            # 以下程式碼有假設：每組漢字之結尾，必有標點符號
            sheet.cells(row_hanzi, col+1).value = "=CHAR(10)"

            # 更新下一組漢字及TLPA標音之位置
            row_hanzi += 4      # 漢字位置
            row_tlpa = row_hanzi + piau_im_row   # TLPA位置: -1 ==> 自動標音； -2 ==> 人工標音

    # 填入文章終止符號：φ
    sheet.cells(row_hanzi-4, 4).value = "φ"
    logging.info(f"已將漢字及TLPA注音填入【{sheet_name}】工作表！")

# =========================================================================
# 主作業程序
# =========================================================================
def main():
    # 檢查是否有指定檔案名稱，若無則使用預設檔名
    filename = sys.argv[1] if len(sys.argv) > 1 else "tmp.txt"
    # 檢查是否有 'ho' 參數，若有則使用標音格式二：【聲母】+【韻母】+【調號】
    if "hu" in sys.argv:  # 若命令行參數包含 'bp'，則使用 BP
        use_tiau_ho = False
    else:
        use_tiau_ho = True
    # 以作用中的Excel活頁簿為作業標的
    wb = xw.apps.active.books.active
    if wb is None:
        logging.error("無法找到作用中的Excel活頁簿。")
        return

    fill_hanzi_and_tlpa(wb,
                        filename=filename,
                        use_tiau_ho=use_tiau_ho,
                        start_row=5,
                        piau_im_row=-2) # -1: 自動標音；-2: 人工標音

if __name__ == "__main__":
    main()
