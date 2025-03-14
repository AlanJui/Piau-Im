# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import re
import unicodedata

import xlwings as xw

# =========================================================================
# 程式區域函式
# =========================================================================
# 用途：從純文字檔案讀取資料並回傳 [(漢字, TLPA), ...] 之格式

def read_text_with_tlpa(filename):
    text_with_tlpa = []
    with open(filename, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip() and not line.startswith('zh.wikipedia.org')]
    for i in range(0, len(lines), 2):
        text_with_tlpa.append((lines[i], lines[i + 1]))
    return text_with_tlpa

# =========================================================================
# 用途：檢查是否為漢字
# =========================================================================
def is_hanzi(char):
    return 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(char, '')

# =========================================================================
# 用途：清理TLPA標音，去除標點符號但保留數字聲調
# =========================================================================
def clean_tlpa(word):
    return re.sub(r'[^a-zA-Z0-9̀-ͯ]', '', word)  # 允許數字和聲調符號

# =========================================================================
# 用途：將漢字及TLPA標音填入Excel指定工作表
# =========================================================================
def fill_hanzi_and_tlpa(wb, filename='tmp.txt', sheet_name='漢字注音', start_row=5):
    sheet = wb.sheets[sheet_name]
    sheet.activate()
    sheet.range('A1').select()

    text_with_tlpa = read_text_with_tlpa(filename)

    for idx, (hanzi, tlpa) in enumerate(text_with_tlpa):
        row_hanzi = start_row + idx * 4      # 漢字位置
        row_tlpa = row_hanzi - 1             # TLPA位置

        # 漢字逐字填入（從D欄開始）
        for col_idx, char in enumerate(hanzi):
            col = 4 + col_idx  # D欄是第4欄
            sheet.cells(row_hanzi, col).value = char
            sheet.cells(row_hanzi, col).select()  # 每字填入後選取以便畫面滾動

        # TLPA逐詞填入（從D欄開始），檢查下方儲存格是否為漢字
        tlpa_words = []
        for word in tlpa.split():
            cleaned_word = clean_tlpa(word)  # 去除標點符號但保留聲調數字
            if '-' in cleaned_word:
                tlpa_words.extend(cleaned_word.split('-'))
            else:
                tlpa_words.append(cleaned_word)

        col = 4
        word_idx = 0

        while word_idx < len(tlpa_words):
            cell_char = sheet.cells(row_hanzi, col).value
            if cell_char and is_hanzi(cell_char):
                sheet.cells(row_tlpa, col).value = tlpa_words[word_idx]
                word_idx += 1
                logging.info(f"已完成填入: {cell_char} - {tlpa_words[word_idx-1]}")
            col += 1

    logging.info(f"已將漢字及TLPA注音填入【{sheet_name}】工作表！")

# =========================================================================
# 主作業程序
# =========================================================================
def main():
    wb = xw.apps.active.books.active
    if wb is None:
        logging.error("無法找到作用中的Excel活頁簿。")
        return

    fill_hanzi_and_tlpa(wb)

if __name__ == "__main__":
    main()
