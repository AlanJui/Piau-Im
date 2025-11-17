# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import re
import sys
import unicodedata

import xlwings as xw

# =========================================================================
# 設定標點符號過濾
# =========================================================================
PUNCTUATIONS = (",", ".", "?", "!")

# TLPA → BP 拼音轉換對應表
TLPA_TO_BP = {
    "eng": "ing",  # 例: teng → ting
    "oa": "ua",  # 例: loan → luan
    "oe": "ue",  # 例: koe → kue
    "chh": "c",  # 例: chhia → cia
    "ch": "z",  # 例: chai → zai
    "oo": "oo",  # BP 保持不變
}

# TLPA → BP 聲調轉換
TLPA_TONE_TO_BP_TONE = {
    "2": "5",
    "3": "3",
    "5": "2",
    "6": "6",
    "7": "7",
    "8": "8",
}

# TLPA 聲調符號 → BP 聲調符號
# "i": TLPA 陰平 (1) → "ī": 閩拼 陰平 (1)
# "í": TLPA 陰上 (2) → "ǐ": 閩拼 上聲 (3)
# "ì": TLPA 陰去 (3) → "ì": 閩拼 陰去 (5)
# "i": TLPA 陰入 (4) → "ī": 閩拼 陰入 (7)
# "î": TLPA 陽平 (5) → "í": 閩拼 陽平 (2)
# "ǐ": TLPA 陽上 (6) → "ǐ": 閩拼 上聲 (4)
# "ī": TLPA 陽去 (7) → "î": 閩拼 陽去 (6)
# "i̍": TLPA 陽入 (8) → "í": 閩拼 陽入 (8)
TLPA_TONE_SYMBOL_TO_BP = {
    # i 部分
    "i": "ī",  # TLPA 陰平 (1) → 閩拼 陰平 (1)
    "í": "ǐ",  # TLPA 陰上 (2) → 閩拼 上聲 (3)
    "ì": "ì",  # TLPA 陰去 (3) → 閩拼 陰去 (5)
    "i": "ī",  # TLPA 陰入 (4) → 閩拼 陰入 (7)
    "î": "í",  # TLPA 陽平 (5) → 閩拼 陽平 (2)
    "ǐ": "ǐ",  # TLPA 陽上 (6) → 閩拼 上聲 (4)
    "ī": "î",  # TLPA 陽去 (7) → 閩拼 陽去 (6)
    "i̍": "í",  # TLPA 陽入 (8) → 閩拼 陽入 (8)

    # u 部分
    "u": "ū",  # TLPA 陰平 (1) → 閩拼 陰平 (1)
    "ú": "ǔ",  # TLPA 陰上 (2) → 閩拼 上聲 (3)
    "ù": "ù",  # TLPA 陰去 (3) → 閩拼 陰去 (5)
    "u": "ū",  # TLPA 陰入 (4) → 閩拼 陰入 (7)
    "û": "ú",  # TLPA 陽平 (5) → 閩拼 陽平 (2)
    "ǔ": "ǔ",  # TLPA 陽上 (6) → 閩拼 上聲 (4)
    "ū": "û",  # TLPA 陽去 (7) → 閩拼 陽去 (6)
    "u̍": "ú",  # TLPA 陽入 (8) → 閩拼 陽入 (8)
}
# =========================================================================
# 程式區域函式
# =========================================================================
# 用途：從純文字檔案讀取資料並回傳 [(漢字, 拼音), ...] 之格式

def read_text_with_tlpa(filename):
    text_with_tlpa = []
    with open(filename, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip() and not line.startswith('zh.wikipedia.org')]
    for i in range(0, len(lines), 2):
        hanzi = lines[i]
        tlpa = lines[i + 1].replace("-", " ")  # 替換 "-" 為空白字元
        text_with_tlpa.append((hanzi, tlpa))
    return text_with_tlpa

# 用途：檢查是否為漢字
def is_hanzi(char):
    return 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(char, '')

# =========================================================================
# 用途：移除標點符號並轉換拼音格式（TLPA 或 BP）
# =========================================================================
def clean_pinyin(word, use_bp=False):
    word = ''.join(ch for ch in word if ch not in PUNCTUATIONS)  # 移除標點符號
    if use_bp:
        for tlpa, bp in TLPA_TO_BP.items():
            word = word.replace(tlpa, bp)  # 轉換為 BP
        # 處理鼻化音標記（TLPA: ann → BP: nna, oann → BP: nuan）
        word = re.sub(r'([iu]?(ai|au|[iuaoe]))nn$', r'n\1', word)
        # 處理以母音開頭但帶有聲調的拼音（í → yí, ú → wú）
        word = ''.join(TLPA_TONE_SYMBOL_TO_BP.get(ch, ch) for ch in word)  # 轉換 TLPA 聲調符號至 BP
        if re.match(r'^[aeiou]', word):
            word = ('y' if word.startswith('i') else 'w') + word
        # 轉換 TLPA 數字聲調至 BP
        word = re.sub(r'(\d)$', lambda m: TLPA_TONE_TO_BP_TONE.get(m.group(1), m.group(1)), word)
    return word

# =========================================================================
# 用途：將漢字及拼音填入Excel指定工作表
# =========================================================================
def fill_hanzi_and_pinyin(wb, filename, use_bp=False, sheet_name='漢字注音', start_row=5):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    text_with_pinyin = read_text_with_tlpa(filename)

    for idx, (hanzi, pinyin) in enumerate(text_with_pinyin):
        row_hanzi = start_row + idx * 4      # 漢字位置
        row_pinyin = row_hanzi - 1           # 拼音位置

        # 漢字逐字填入（從D欄開始）
        for col_idx, char in enumerate(hanzi):
            col = 4 + col_idx  # D欄是第4欄
            sheet.cells(row_hanzi, col).value = char
            sheet.cells(row_hanzi, col).select()  # 每字填入後選取以便畫面滾動

        # 拼音逐詞填入（從D欄開始），檢查下方儲存格是否為漢字
        pinyin_words = [clean_pinyin(word, use_bp) for word in pinyin.split()]
        col = 4
        word_idx = 0

        while word_idx < len(pinyin_words):
            cell_char = sheet.cells(row_hanzi, col).value
            if cell_char and is_hanzi(cell_char):
                sheet.cells(row_pinyin, col).value = pinyin_words[word_idx]
                word_idx += 1
                print(f"({row_hanzi}, {col}) 已填入: {cell_char} - {pinyin_words[word_idx-1]}")
            col += 1

    logging.info(f"已將漢字及拼音填入【{sheet_name}】工作表！")

# =========================================================================
# 主作業程序
# =========================================================================
def main():
    filename = sys.argv[1] if len(sys.argv) > 1 else "tmp.txt"
    use_bp = "bp" in sys.argv  # 若命令行參數包含 'bp'，則使用 BP

    wb = xw.apps.active.books.active
    if wb is None:
        logging.error("無法找到作用中的Excel活頁簿。")
        return

    fill_hanzi_and_pinyin(wb, filename, use_bp)

if __name__ == "__main__":
    main()
