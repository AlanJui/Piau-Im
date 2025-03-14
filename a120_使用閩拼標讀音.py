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
    "oa": "ua",  # 例: loan → luan
    "chh": "c",  # 例: chhia → cia
    "ch": "z",  # 例: chai → zai
    "oo": "oo",  # BP 保持不變
}

# TLPA → BP 聲調轉換
TLPA_TONE_TO_BP_TONE = {
    "7": "6",
    "2": "3",
    "3": "5",
    "5": "2",
    "6": "4",
    "4": "7",
    "8": "8"
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
        if re.match(r'^[íìîïi]', word, re.IGNORECASE):
            word = 'y' + word
        elif re.match(r'^[úùûüu]', word, re.IGNORECASE):
            word = 'w' + word
        # 轉換 TLPA 聲調至 BP
        word = re.sub(r'(\d)$', lambda m: TLPA_TONE_TO_BP_TONE.get(m.group(1), m.group(1)), word)
    return word
    íìîǐīi̍ # 2：陰上、3：陰去、5：陽平、6：陽上、7：陽去、8：陽入
    ǎāàāáâá
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
                logging.info(f"已完成填入: {cell_char} - {pinyin_words[word_idx-1]}")
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
