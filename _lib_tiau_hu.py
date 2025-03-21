import unicodedata


def tones_to_hex(tones):
    """將 tones 字串轉換成 16 進制數值"""
    # return [hex(ord(c)) for c in tones]
    list =  [hex(ord(c)) for c in tones]
    string = ''.join(list)
    return string

def tones_to_unicode_format(tones):
    # """將 tones 轉換成 \u0xxx 格式"""
    # return [f"\\u{ord(c):04x}" for c in tones]
    return [f"\\u{ord(c):04x}" for c in tones]

# =========================================================================
# 【帶調符拼音】轉【帶調號拼音】
# =========================================================================
# 設定標點符號過濾
PUNCTUATIONS = (",", ".", "?", "!", ":", ";", "\u200B")

# 調符對映表（帶調符之元音/韻化輔音 → 不帶調符之拼音字母、調號數值）
tiau_hu_mapping = {
    "á": ("a", "2"), "à": ("a", "3"), "â": ("a", "5"), "ǎ": ("a", "6"), "ā": ("a", "7"), "a̍": ("a", "8"), "a̋": ("a", "9"),
    "Á": ("A", "2"), "À": ("A", "3"), "Â": ("A", "5"), "Ǎ": ("A", "6"), "Ā": ("A", "7"), "A̍": ("A", "8"), "A̋": ("A", "9"),
    "é": ("e", "2"), "è": ("e", "3"), "ê": ("e", "5"), "ě": ("e", "6"), "ē": ("e", "7"), "e̍": ("e", "8"), "e̋": ("e", "9"),
    "É": ("E", "2"), "È": ("E", "3"), "Ê": ("E", "5"), "Ě": ("E", "6"), "Ē": ("E", "7"), "E̍": ("E", "8"), "E̋": ("E", "9"),
    "í": ("i", "2"), "ì": ("i", "3"), "î": ("i", "5"), "ǐ": ("i", "6"), "ī": ("i", "7"), "i̍": ("i", "8"), "i̋": ("i", "9"),
    "Í": ("I", "2"), "Ì": ("I", "3"), "Î": ("I", "5"), "Ǐ": ("I", "6"), "Ī": ("I", "7"), "I̍": ("I", "8"), "I̋": ("I", "9"),
    "ó": ("o", "2"), "ò": ("o", "3"), "ô": ("o", "5"), "ǒ": ("o", "6"), "ō": ("o", "7"), "o̍": ("o", "8"), "ő": ("o", "9"),
    "Ó": ("O", "2"), "Ò": ("O", "3"), "Ô": ("O", "5"), "Ǒ": ("O", "6"), "Ō": ("O", "7"), "O̍": ("O", "8"), "Ő": ("O", "9"),
    "ú": ("u", "2"), "ù": ("u", "3"), "û": ("u", "5"), "ǔ": ("u", "6"), "ū": ("u", "7"), "u̍": ("u", "8"), "ű": ("u", "9"),
    "Ú": ("U", "2"), "Ù": ("U", "3"), "Û": ("U", "5"), "Ǔ": ("U", "6"), "Ū": ("U", "7"), "U̍": ("U", "8"), "Ű": ("U", "9"),
    "ḿ": ("m", "2"), "m̀": ("m", "3"), "m̂": ("m", "5"), "m̌": ("m", "6"), "m̄": ("m", "7"), "m̍": ("m", "8"), "m̋": ("m", "9"),
    "Ḿ": ("M", "2"), "M̀": ("M", "3"), "M̂": ("M", "5"), "M̌": ("M", "6"), "M̄": ("M", "7"), "M̍": ("M", "8"), "M̋": ("M", "9"),
    "ń": ("n", "2"), "ǹ": ("n", "3"), "n̂": ("n", "5"), "ň": ("n", "6"), "n̄": ("n", "7"), "n̍": ("n", "8"), "n̋": ("n", "9"),
    "Ń": ("N", "2"), "Ǹ": ("N", "3"), "N̂": ("N", "5"), "Ň": ("N", "6"), "N̄": ("N", "7"), "N̍": ("N", "8"), "N̋": ("N", "9"),
}

def tng_tiau_ho(im_piau: str, kan_hua: bool = False) -> str:
    """
    將帶聲調符號的台語音標轉換為不帶聲調符號的台語音標（音標 + 調號）
    :param im_piau: str - 台語音標輸入
    :param kan_hua: bool - 簡化：若是【簡化】，聲調值為 1 或 4 ，去除調號值
    :return: str - 轉換後的台語音標
    """
    # 遇標點符號，不做轉換處理，直接回傳
    if im_piau[-1] in PUNCTUATIONS:
        return im_piau

    # 若【音標】末端為數值，表音標已是【帶調號拼音】，直接回傳
    u_tiau_ho = True if im_piau[-1] in "123456789" else False
    if u_tiau_ho: return im_piau

    # 將傳入【音標】字串，以標準化之 NFC 組合格式，調整【帶調符拼音字母】；
    # 令以下之處理作業，不會發生【看似相同】的【帶調符拼音字母】，其實使用
    # 不同之 Unicode 編碼
    im_piau = unicodedata.normalize("NFC", im_piau)

    #--------------------------------------------------------------------------------
    # 以【元音及韻化輔音清單】，比對傳入之【音標】，找出對應之【基本拼音字母】與【調號】
    #--------------------------------------------------------------------------------
    tone_number = "1"  # 初始化調號為 1
    number = "1"  # 明確初始化 number 變數，以免未設定而發生錯誤
    for tone_mark, (base_char, number) in tiau_hu_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除聲調符號，保留基本母音
            break
    else:
        number = "1"  # 若沒有任何調符，number強制為1

    # 依是否要【簡化】之設定，處理【調號】1 或 4 是否要略去
    if kan_hua and number in ["1", "4"]:
        # 若是【簡化】，且聲調值為 1 或 4 ，去除調號值
        tone_number = ""
    else:
        # 若未要求【簡化】，聲調值置於【音標】末端
        if number not in ["1", "4"]:
            tone_number = number
        else:
            if im_piau[-1] in "hptk":
                # 【音標】末端為【hptk】之一，則為【陰入調】，聲調值為 4
                tone_number = "4"
            else:
                tone_number = "1"

    return im_piau + tone_number

def ut01(test_cases: dict):
    for key, value in test_cases.items():
        # print(f"漢字: {key}, 白話音標: {value}")
        print(f"{key} [ {value} ]")

def ut02(test_cases: dict):
    # ut01(test_cases)

    kan_hua = False
    for key, im_piau in test_cases.items():
        # print(f"漢字: {key}, 白話音標: {value}")
        print(f"{key} [ {im_piau} ]")
        # 以【元音及韻化輔音清單】，比對傳入之【音標】，找出對應之【基本拼音字母】與【調號】
        tone_number = "1"
        for tone_mark, (base_char, number) in tiau_hu_mapping.items():
            if tone_mark in im_piau:
                im_piau = im_piau.replace(tone_mark, base_char)  # 移除聲調符號，保留基本母音
                break

        # 依是否要【簡化】之設定，處理【調號】1 或 4 是否要略去
        if kan_hua and number in ["1", "4"]:
            # 若是【簡化】，且聲調值為 1 或 4 ，去除調號值
            tone_number = ""
        else:
            # 若未要求【簡化】，聲調值置於【音標】末端
            if number not in ["1", "4"]:
                tone_number = number
            else:
                if im_piau[-1] in "hptk":
                    # 【音標】末端為【hptk】之一，則為【陰入調】，聲調值為 4
                    tone_number = "4"
                else:
                    tone_number = "1"

        im_piau_tiau_ho = im_piau + tone_number
        print(f"im_piau_tiau_ho = {im_piau_tiau_ho}")

if __name__ == "__main__":
    # test_cases = {
    #     "園": 'ôan',
    #     "為": 'ûi',
    #     "獨": 'to̍k',
    #     "悟": 'ngōo',
    #     "途": 'tôo',
    #     "遠": 'oán',
    #     "昨": 'cha̍k',
    # }
    # ut01(test_cases)

    # test_cases = {
    #     "東": 'tong',
    # }
    # for key, im_piau in test_cases.items():
    #     # print(f"漢字: {key}, 白話音標: {value}")
    #     print(f"{key} [ {im_piau} ]")
    #     # 轉換【帶調符拼音】為【帶調號拼音】
    #     im_piau_tiau_ho = tng_tiau_ho(im_piau, kan_hua=False)
    #     print(f"im_piau_tiau_ho = {im_piau_tiau_ho}")

    test_cases = {
        "泉": 'tsuann5',
        "涓": 'kuan1',
        "而": 'ji5',
        "始": 'si2',
        "流": 'lau5',
    }

    for key, im_piau in test_cases.items():
        print(f"{key} [ {im_piau} ]")
        # 轉換【帶調符拼音】為【帶調號拼音】
        # u_tiau_ho = True if im_piau[-1] in "123456789" else False
        # print(f"u_tiau_ho = {u_tiau_ho}")
        im_piau_tiau_ho = tng_tiau_ho(im_piau, kan_hua=False)
        print(f"im_piau_tiau_ho = {im_piau_tiau_ho}")