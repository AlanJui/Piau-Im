# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import sys
import unicodedata

import xlwings as xw

# =========================================================================
# 設定標點符號過濾
# =========================================================================
PUNCTUATIONS = (",", ".", "?", "!", ":")

# =========================================================================
# 程式區域函式
# =========================================================================
# 用途：從純文字檔案讀取資料並回傳 [(漢字, TLPA), ...] 之格式

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
# 用途：移除標點符號並轉換TLPA+拼音格式
# =========================================================================
def clean_tlpa(word):
    word = ''.join(ch for ch in word if ch not in PUNCTUATIONS)  # 移除標點符號
    word = word.replace("oa", "ua")  # TLPA+ 調整，將 "oa" 變為 "ua"
    if word.startswith("chh"):
        word = "c" + word[3:]
    elif word.startswith("ch"):
        word = "z" + word[2:]
    return word

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
        tlpa_words = [clean_tlpa(word) for word in tlpa.split()]
        col = 4
        word_idx = 0

        while word_idx < len(tlpa_words):
            cell_char = sheet.cells(row_hanzi, col).value
            if cell_char and is_hanzi(cell_char):
                sheet.cells(row_tlpa, col).value = tlpa_words[word_idx]
                word_idx += 1
                print(f"（{row_tlpa}, {col}）已填入: {cell_char} - {tlpa_words[word_idx-1]}")
            col += 1

    logging.info(f"已將漢字及TLPA注音填入【{sheet_name}】工作表！")

# =========================================================================
# 主作業程序
# =========================================================================
def main():
    # 檢查是否有指定檔案名稱，若無則使用預設檔名
    filename = sys.argv[1] if len(sys.argv) > 1 else "tmp.txt"
    # 以作用中的Excel活頁簿為作業標的
    wb = xw.apps.active.books.active
    if wb is None:
        logging.error("無法找到作用中的Excel活頁簿。")
        return

    fill_hanzi_and_tlpa(wb)

if __name__ == "__main__":
    main()
