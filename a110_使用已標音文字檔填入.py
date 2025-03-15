import logging
import re
import sys
import unicodedata

import xlwings as xw

# TLPA 聲調符號對應數值
# fmt: off
TONE_MAP = {
    "á": "2", "à": "3", "a̍": "8", "â": "5", "ǎ": "6", "ā": "7",  # a
    "é": "2", "è": "3", "e̍": "8", "ê": "5", "ě": "6", "ē": "7",  # e
    "í": "2", "ì": "3", "i̍": "8", "î": "5", "ǐ": "6", "ī": "7",  # i
    "ó": "2", "ò": "3", "o̍": "8", "ô": "5", "ǒ": "6", "ō": "7",  # o
    "ú": "2", "ù": "3", "u̍": "8", "û": "5", "ǔ": "6", "ū": "7",  # u
    "ń": "2", "ň": "6", "ñ": "5"  # 特殊鼻音
}
# fmt: on

# 需過濾的標點符號和控制字元
# PUNCTUATION = [',', '.', '!', '?', ':', '\u200b']
PUNCTUATION = [',', '.', '!', '?', ':', '​']

# 用途：將 TLPA 拼音中的聲調符號轉換為數字
def convert_tlpa_tone(tlpa_word):
    tone = "1"  # 預設為陰平調
    plain_word = ""

    for char in tlpa_word:
        if char in TONE_MAP:
            tone = TONE_MAP[char]  # 取得對應的聲調數值
            plain_word += unicodedata.normalize("NFD", char)[0]  # 去掉聲調符號
        else:
            plain_word += char

    # 若尾碼為 h/p/t/k，則屬於陰入調（4調）
    if plain_word[-1] in "hptk":
        tone = "4"

    return plain_word + tone


# 用途：讀取文字檔案並轉換 TLPA 拼音
def read_text_with_tlpa(filename):
    text_with_tlpa = []
    with open(filename, "r", encoding="utf-8") as f:
        lines = [line.strip().replace("\u200b", "") for line in f if line.strip() and not line.startswith("zh.wikipedia.org")]

    for i in range(0, len(lines), 2):
        hanzi = lines[i]
        tlpa = lines[i + 1].replace("-", " ")  # 替換 "-" 為空白字元

        # 使用正則表達式確保標點符號成為獨立詞
        tlpa_words = re.findall(r'\w+|[{}]'.format(re.escape(''.join(PUNCTUATION))), tlpa)

        text_with_tlpa.append((hanzi, tlpa_words))

    return text_with_tlpa


# 用途：將漢字及 TLPA 注音填入 Excel 指定工作表
def fill_hanzi_and_tlpa(wb, filename="tmp.txt", sheet_name="漢字注音", start_row=5):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    text_with_tlpa = read_text_with_tlpa(filename)

    for idx, (hanzi, tlpa_list) in enumerate(text_with_tlpa):
        row_hanzi = start_row + idx * 4  # 漢字位置
        row_tlpa = row_hanzi - 1  # TLPA 位置

        col = 4  # Excel 欄位從 D 欄開始
        word_idx = 0

        while word_idx < len(tlpa_list):
            if col - 4 < len(hanzi):  # 確保不超出漢字範圍
                char = hanzi[col - 4]
                sheet.cells(row_hanzi, col).value = char  # 填入漢字

                tlpa = tlpa_list[word_idx]
                if tlpa in PUNCTUATION:
                    # 如果 TLPA 是標點符號，將其填入漢字行，並繼續處理
                    sheet.cells(row_tlpa, col).value = ""
                else:
                    sheet.cells(row_tlpa, col).value = convert_tlpa_tone(tlpa)  # 轉換並填入拼音

                print(f"（{row_tlpa}, {col}）已填入: {char} - {convert_tlpa_tone(tlpa)}")
                word_idx += 1

            col += 1

    logging.info(f"已將漢字及 TLPA 注音填入【{sheet_name}】工作表！")


# 主作業程序
def main():
    filename = sys.argv[1] if len(sys.argv) > 1 else "tmp.txt"
    wb = xw.apps.active.books.active
    if wb is None:
        logging.error("無法找到作用中的 Excel 活頁簿。")
        return
    fill_hanzi_and_tlpa(wb, filename)


if __name__ == "__main__":
    main()
