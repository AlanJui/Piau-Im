# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import re
import sys
import unicodedata

import xlwings as xw

# =========================================================
# 解構音標 = 聲母 + 韻母 + 調號
# =========================================================

# 上標數字與普通數字的映射字典
superscript_digit_mapping = {
    '⁰': '0',
    '¹': '1',
    '²': '2',
    '³': '3',
    '⁴': '4',
    '⁵': '5',
    '⁶': '6',
    '⁷': '7',
    '⁸': '8',
    '⁹': '9',
}

def replace_superscript_digits(input_str):
    return ''.join(superscript_digit_mapping.get(char, char) for char in input_str)


# =========================================================================
# 設定標點符號過濾
# =========================================================================
# PUNCTUATIONS2 = (",", ".", "?", "!", ":", ";")
PUNCTUATIONS = (",", ".", "?", "!", ":", ";", "\u200B")

# 確認音標的拼音字母中不帶：標點符號、控制字元
def clean_im_piau(im_piau: str) -> str:
    # 移除標點符號
    im_piau = ''.join(ji_bu for ji_bu in im_piau if ji_bu not in PUNCTUATIONS)
    # 透過正規化的 Unicode 標準分解 NFD，拆解聲調符號
    im_piau = unicodedata.normalize("NFD", im_piau)

    su_ji = im_piau[0]  # 保存第一個字母
    im_piau = im_piau.lower()  # 轉為小寫

    # **新增鼻音處理：將 ⁿ（U+207F）轉換為 nn**
    im_piau = im_piau.replace("ⁿ", "nn")
    im_piau = im_piau.replace("hⁿ", "nnh")
    # im_piau = im_piau.replace("o͘", "oo")  # 替換 o͘ (o + 鼻音符號)

    # 替換音標變化
    # im_piau = re.sub(r"u[\u0300\u0301\u0302\u0304\u030D]?e", "ui", im_piau)
    im_piau = re.sub(r"o[\u0300\u0301\u0302\u0304\u030D]?a", "ua", im_piau)
    im_piau = re.sub(r"o[\u0300\u0301\u0302\u0304\u030D]?e", "ue", im_piau)
    im_piau = re.sub(r"e[\u0300\u0301\u0302\u0304\u030D]?ng", "ing", im_piau)
    im_piau = re.sub(r"e[\u0300\u0301\u0302\u0304\u030D]?k", "ik", im_piau)

    #-------------------------------------------------------------------------
    #
    #-------------------------------------------------------------------------
    # 聲調符號對應調值的映射
    tone_mapping = {
        "\u0300": "3",  # 陰去 ò
        "\u0301": "2",  # 陰上 ó
        "\u0302": "5",  # 陽平 ô
        "\u0304": "7",  # 陽去 ō
        "\u0306": "9",  # 輕声 ŏ
        "\u030C": "6",  # 陽上 ǒ
        "\u030D": "8",  # 陽入 o̍h
    }

    # 替換白話字母為oo，並附加聲調號
    # 找到帶鼻化符號(͘)的 o 或 ô，將其轉成對應的帶調符號 + o
    im_piau = re.sub(
        r"([aeiou])([\u0300\u0301\u0302\u0304\u030D])?\u0358",
        lambda m: f"{m.group(1)}{m.group(2) if m.group(2) else ''}o",
        im_piau
    )

    if im_piau.startswith("chh"):
        im_piau = "c" + im_piau[3:]
    elif im_piau.startswith("ch"):
        im_piau = "z" + im_piau[2:]

    if su_ji.isupper():
        im_piau = im_piau.capitalize()

    im_piau = unicodedata.normalize("NFC", im_piau)  # 重新組合聲調符號（標準組合 NFC）
    return im_piau


# =========================================================================
# 將使用聲調符號的 TLPA 拼音轉為改用調號數值的 TLPA 拼音
# =========================================================================

# 聲調符號對應表（帶調號母音 → 對應數字）
tone_mapping = {
    "A̍": ("A", "8"),
    "Á": ("A", "2"),
    "Ǎ": ("A", "6"),
    "Â": ("A", "5"),
    "Ā": ("A", "7"),
    "À": ("A", "3"),
    "A̋": ("A", "9"),

    "E̍": ("E", "8"),
    "É": ("E", "2"),
    "Ě": ("E", "6"),
    "Ê": ("E", "5"),
    "Ē": ("E", "7"),
    "È": ("E", "3"),
    "E̋": ("E", "9"),

    "I̍": ("I", "8"),
    "Í": ("I", "2"),
    "Î": ("I", "5"),
    "Ǐ": ("I", "6"),
    "Ī": ("I", "7"),
    "Ì": ("I", "3"),
    "I̋": ("I", "9"),

    "O̍": ("O", "8"),
    "Ó": ("O", "2"),
    "Ǒ": ("O", "6"),
    "Ô": ("O", "5"),
    "Ō": ("O", "7"),
    "Ò": ("O", "3"),
    "Ő ": ("O", "9"),

    "U̍": ("U", "8"),
    "Ú": ("U", "2"),
    "Ǔ": ("U", "6"),
    "Û": ("U", "5"),
    "Ū": ("U", "7"),
    "Ù": ("U", "3"),
    "Ű ": ("U", "9"),

    "M̍": ("M", "8"),
    "Ḿ": ("M", "2"),
    "M̌": ("M", "6"),
    "M̂": ("M", "5"),
    "M̄": ("M", "7"),
    "M̀": ("M", "3"),
    "M̋": ("M", "9"),

    "N̍": ("N", "8"),
    "Ń": ("N", "2"),
    "Ň": ("N", "6"),
    "N̂": ("N", "5"),
    "N̄": ("N", "7"),
    "Ǹ": ("N", "3"),
    "N̋": ("N", "9"),

    "a̍": ("a", "8"),
    "á": ("a", "2"),
    "ǎ": ("a", "6"),
    "â": ("a", "5"),
    "ā": ("a", "7"),
    "à": ("a", "3"),
    "a̋": ("a", "9"),

    "e̍": ("e", "8"),
    "é": ("e", "2"),
    "ě": ("e", "6"),
    "ê": ("e", "5"),
    "ē": ("e", "7"),
    "è": ("e", "3"),
    "e̋": ("e", "9"),

    "i̍": ("i", "8"),
    "í": ("i", "2"),
    "ǐ": ("i", "6"),
    "î": ("i", "5"),
    "ī": ("i", "7"),
    "ì": ("i", "3"),
    "i̋": ("i", "9"),

    "o̍": ("o", "8"),
    "ó": ("o", "2"),
    "ǒ": ("o", "6"),
    "ô": ("o", "5"),
    "ō": ("o", "7"),
    "ò": ("o", "3"),
    "ő ": ("o", "9"),

    "u̍": ("u", "8"),
    "ú": ("u", "2"),
    "ǔ": ("u", "6"),
    "û": ("u", "5"),
    "ū": ("u", "7"),
    "ù": ("u", "3"),
    "ű ": ("u", "9"),

    "m̍": ("m", "8"),
    "ḿ": ("m", "2"),
    "m̌": ("m", "6"),
    "m̂": ("m", "5"),
    "m̄": ("m", "7"),
    "m̀": ("m", "3"),
    "m̋": ("m", "9"),

    "n̍": ("n", "8"),
    "ń": ("n", "2"),
    "ň": ("n", "6"),
    "n̂": ("n", "5"),
    "n̄": ("n", "7"),
    "ǹ": ("n", "3"),
    "n̋": ("n", "9"),
}

# 韻母轉換字典
un_bu_tng_huan_map_dict = {
    # 'onn': 'oonn',      # 雅俗通十五音：扛
    'ueinn': 'uenn',    # 雅俗通十五音：檜
    'uei': 'ue',        # 雅俗通十五音：檜
    'ue': 'ui',
    'ereh': 'ueh',      # ereh = [əeh]
    'erh': 'eh',        # er（ㄜ）= [ə]
    'ere': 'ue',        # ere = [əe]
    'er': 'e',          # er（ㄜ）= [ə]
    'ee': 'e',          # ee（ㄝ）= [ɛ]
    'or': 'o',          # or（ㄜ）= [ə]
    'ir': 'i',          # ir（ㆨ）= [ɯ] / [ɨ]
    'eng': 'ing',       # 白話字：eng ==> 閩南語：ing
    'oa': 'ua',         # 白話字：oa ==> 閩南語：ua
    'oe': 'ue',         # 白話字：oe ==> 閩南語：ue
    'ei': 'e',          # 雅俗通十五音：稽
    'ou': 'oo',         # 雅俗通十五音：沽
    'ur': 'u',          # 雅俗通十五音：艍
    'ek': 'ik',
    'ⁿ' : 'nn',
}


# =========================================================================
# 將帶聲調符號的【拼音】轉為用數值表示的【拼音】
# =========================================================================

def tng_tiau_ho(im_piau: str) -> str:
    """
    將帶聲調符號的台語音標轉換為不帶聲調符號的台語音標（音標 + 調號）
    :param im_piau: str - 台語音標輸入
    :return: str - 轉換後的台語音標
    """
    # **遍歷所有可能的聲調符號**
    tone_number = "0"
    for tone_mark, (base_char, number) in tone_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除聲調符號，保留基本母音
            tone_number = number
            break

    # print(f"im_piau + tone_number = {im_piau + tone_number}")
    return im_piau + tone_number

def tiau_hu_tng_tiau_ho(im_piau: str) -> str:
    """
    將帶聲調符號的台語音標轉換為不帶聲調符號的台語音標（音標 + 調號）
    :param im_piau: str - 台語音標輸入
    :return: str - 轉換後的台語音標
    """
    # **重要**：先將字串標準化為 NFC 格式，統一處理 Unicode 差異
    im_piau = unicodedata.normalize("NFC", im_piau)

    # 1. 先處理聲調轉換
    tone_number = ""
    for tone_mark, (base_char, number) in tone_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除調號，還原原始母音
            tone_number = number  # 記錄對應的聲調數字
            break  # 只會有一個聲調符號，找到就停止

    # 2. 若有聲調數字，則加到末尾
    if tone_number:
        return im_piau + tone_number

    return im_piau  # 若無聲調符號則不變更


# =========================================================
# 韻母轉換
# =========================================================
def un_bu_tng_huan(un_bu: str) -> str:
    """
    將輸入的韻母依照轉換字典進行轉換
    :param un_bu: str - 韻母輸入
    :return: str - 轉換後的韻母結果
    """
    # **新增鼻音處理：將 ⁿ（U+207F）轉換為 nn**
    un_bu = un_bu.replace("ⁿ", "nn")

    # 韻母轉換，若不存在於字典中則返回原始韻母
    return un_bu_tng_huan_map_dict.get(un_bu, un_bu)

def split_tai_gi_im_piau(im_piau: str) -> list:
    # 將輸入的台語音標轉換為小寫
    im_piau = im_piau.lower()
    # 查檢【台語音標】是否符合【標準】=【聲母】+【韻母】+【調號】
    tiau = im_piau[-1]
    tiau = replace_superscript_digits(str(tiau))

    # 矯正未標明陰平/陰入調號的情況
    if tiau in ['p', 't', 'k', 'h']:
        tiau = '4'
        im_piau += tiau
    elif tiau in ['a', 'e', 'i', 'o', 'u', 'm', 'n', 'g']:
        tiau = '1'
        im_piau += tiau

    # 聲母相容性轉換
    if im_piau.startswith("tsh"):
        im_piau = im_piau.replace("tsh", "c", 1)
    elif im_piau.startswith("ts"):
        im_piau = im_piau.replace("ts", "z", 1)
    elif im_piau.startswith("chh"):
        im_piau = im_piau.replace("chh", "c", 1)
    elif im_piau.startswith("ch"):
        im_piau = im_piau.replace("ch", "z", 1)

    # 定義聲母的正規表示式，包括常見的聲母，但不包括 m 和 ng
    siann_bu_pattern = re.compile(r"(b|c|z|g|h|j|kh|k|l|m(?!\d)|ng(?!\d)|n|ph|p|s|th|t|Ø)")
    un_bu_as_m_or_ng_pattern = re.compile(r"(m|ng)\d")

    result = []

    # 首先檢查是否是 m 或 ng 當作韻母的特殊情況
    if un_bu_as_m_or_ng_pattern.match(im_piau):
        siann_bu = ""
        un_bu = im_piau[:-1]
        tiau = im_piau[-1]
    else:
        siann_bu_match = siann_bu_pattern.match(im_piau)
        if siann_bu_match:
            siann_bu = siann_bu_match.group()
            un_bu = im_piau[len(siann_bu):-1]
        else:
            siann_bu = ""
            un_bu = im_piau[:-1]

    # 轉換韻母
    un_bu = un_bu_tng_huan(un_bu)

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result

def clean_tlpa(im_piau: str) -> str:
    org_im_piau = im_piau
    # su_ji = im_piau[0]
    # 去除標點符號、控制字元
    im_piau = clean_im_piau(im_piau)
    # 轉換帶聲調符號的 TLPA 拼音為數值表示的 TLPA 拼音
    # im_piau = tng_tiau_ho(im_piau)
    im_piau = tiau_hu_tng_tiau_ho(im_piau)
    # 轉換 TLPA 音標使用之【聲母】及【韻母】, po_ci: bool = False
    siann_bu, un_bu, tiau = split_tai_gi_im_piau(im_piau=im_piau)
    return f"{siann_bu}{un_bu}{tiau}"


# =========================================================================
# 程式區域函式
# =========================================================================

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
                # else:
                #     # 若讀入之TLPA音標非標點符號，且使用標音格式二，則轉換為【聲母】+【韻母】+【調號】
                #     if use_tiau_ho:
                #         tlpa_word = tiau_hu_tng_tiau_ho(tlpa_word)
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
