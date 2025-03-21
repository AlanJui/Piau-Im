import re
import unicodedata

# 韻母轉換字典
un_bu_mapping = {
    'ee': 'e', 'er': 'e', 'erh': 'eh', 'or': 'o', 'ere': 'ue', 'ereh': 'ueh',
    'ir': 'i', 'eng': 'ing', 'oa': 'ua', 'oe': 'ue', 'oai': 'uai', 'ei': 'e',
    'ou': 'oo', 'onn': 'oonn', 'uei': 'ue', 'ueinn': 'uenn', 'ur': 'u',
    'ⁿ': 'nn',
}

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

def separate_tone(s):
    """拆解帶調字母為無調字母與調號"""
    decomposed = unicodedata.normalize('NFD', s)
    letters = ''.join(c for c in decomposed if unicodedata.category(c) != 'Mn')
    tones = ''.join(c for c in decomposed if unicodedata.category(c) == 'Mn' and c != '\u0358')
    return letters, tones

def apply_tone(im_piau, tone):
    """聲調符號重新加回第一個母音字母上"""
    vowels = 'aeiouAEIOU'
    for i, c in enumerate(im_piau):
        if c in vowels:
            return unicodedata.normalize('NFC', im_piau[:i+1] + tone + im_piau[i+1:])
    return unicodedata.normalize('NFC', im_piau[0] + tone + im_piau[1:])

# 處理 o͘ 韻母特殊情況的函數
def handle_o_dot(im_piau):
    # 依 Unicode 解構標準（NFD）分解傳入之【音標】，取得解構後之【拼音字母與調符】
    decomposed = unicodedata.normalize('NFD', im_piau)
    # 找出 o + 聲調 + 鼻化符號的特殊組合
    match = re.search(r'(o)([\u0300\u0301\u0302\u0304\u030B\u030C\u030D]?)(\u0358)', decomposed, re.I)
    if match:
        # 捕獲【音標】，其【拼音字母】有 o 長音字母，且其右上方帶有圓點調符（\u0358）： o͘
        letter, tone, nasal = match.groups()
        # 將 o 長音字母，轉換成【拼音字母】 oo，再附回聲調
        # replaced = f"{letter}{letter}{tone}"
        replaced = f"{letter}{tone}{letter}"
        # 重組字串
        decomposed = decomposed.replace(match.group(), replaced)
    # 依 Unicode 組合標準（NFC）重構【拼音字母與調符】，組成轉換後之【音標】
    return unicodedata.normalize('NFC', decomposed)


def tng_un_bu(im_piau: str) -> str:
    # 帶調符之白話字韻母 o͘ ，轉換為【帶韻符之 oo 韻母】
    im_piau = handle_o_dot(im_piau)

    # 解構【帶調符音標】，轉成：【無調符音標】、【聲調符號】
    letters, tone = separate_tone(im_piau)

    # 以【無調符音標】，轉換【韻母】
    sorted_keys = sorted(un_bu_mapping, key=len, reverse=True)
    for key in sorted_keys:
        if key in letters:
            letters = letters.replace(key, un_bu_mapping[key])
            break

    if tone:
        letters = apply_tone(letters, tone)

    return letters

# print(separate_tone("á"))  # ('a', '́')
# print(separate_tone("pô"))  # ('po', '̂')

# list = ["á", "pô"]
# for i in list:
#     letters, tones = separate_tone(i)  # ('a', '́') ('po', '̂')
#     # print(f"{letters}{tones_to_hex(tones)}")
#     print(f"{letters}{tones_to_unicode_format(tones)}")

# 測試
# test_cases = ["á", "ê", "ô", "ū", "pô", "ngô͘"]
# test_cases = ["á", "ê", "ô", "ū", "pô"]

# for test in test_cases:
#     letters, tones = separate_tone(test)
#     unicode_tones = tones_to_unicode_format(tones)
#     print(f"Input: {test}, Letters: {letters}, Tones: {tones}, Unicode: {unicode_tones}")

# test_cases = ["ngô͘", "pô", ]
# test_cases = ['Kue', 'kì', 'lâi', 'ê', '!', 'Tiân', 'ôan', 'chiong', 'û', 'hô', 'put', 'kue', '?']
test_cases = ['ziaⁿ5']
for test in test_cases:
    print(tng_un_bu(test), end= " ")  # ngoo
print()