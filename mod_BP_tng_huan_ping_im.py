"""
將【台語音標（TLPA+）】轉換成【閩拚（bp）】。
"""

import re

# 聲母轉換對照表（【索引】字串排序，需由長到短）
SIANN_BU_TNG_UANN_PIAU = {
    "tsh": "c",
    "ts": "z",
    # 二字母
    "ph": "p",  # ㄆ → p (雙唇音/清音：塞音/送氣)
    "th": "t",  # ㄊ → t (齒齦音/清音：塞音/送氣)
    "kh": "k",  # ㄎ → k（軟顎音/清音：塞音/送氣）
    "ng": "ggn", # ㄫ → ng（軟顎音/濁音：鼻音）
    # 一字母
    # 雙唇音
    "p": "b",  # ㄅ → b（雙唇音/清音：塞音不送氣）
    "b": "bb",  # ㆠ → bb（雙唇音/濁音：塞音不送氣）
    "m": "bbn",  # ㄇ → m（雙唇音/濁音：鼻音）
    # ------------------------------
    # 齒齦音
    "t": "d",  # ㄉ → d（齒齦音/清音：塞音/不送氣）
    "n": "ln",  # ㄋ → n（齒齦音/濁音：鼻音）
    "l": "l",  # ㄌ → l（齒齦音/濁音：邊音）
    # ------------------------------
    # 齒齦音
    "z": "z",  # ㄗ → z (齒齦音/清音：塞音/不送氣)
    "j": "zz",  # ㆡ → zz（齒齦音/濁音：塞擦音/不送氣）
    "c": "c",  # ㄘ → c (齒齦音/清音：塞音/送氣)
    "s": "s",  # ㄙ → s（齒齦音/清音：擦音）
    # ------------------------------
    # 軟顎音
    "k": "g",  # ㄍ → g（軟顎音/清音：塞音/不送氣）
    "g": "gg",  # ㆣ → gg（軟顎音/濁音：塞音/不送氣）
    # ------------------------------
    # 聲門音
    "h": "h",  # ㄏ → h（聲門音／擦音：聲門音／清音）
}

# 【齒音聲母+i】轉換對照表
# 【齒音聲母】：TLPA: 舌尖前音/TL: 舌齒音
CI_IM_TNG_UANN_PIAU: dict[str, str] = {
    # "zzi": "jji",  # ㆢ：ji → jj+i
    # "zi": "ji",  # ㄐ：z+i → j+i
    # "ci": "chi",  # ㄑ：c+i → ch+i
    # "si": "shi",  # ㄒ：s+i → sh+i
}

# 韻母轉換對照表（【索引】字串排序，需由長到短）
# （1）複合韻母：
# ai, au/ao, ia, iu, io, ua, ui, ue, iau/iao, uai
# （2）鼻化韻母：
# ann, inn, enn, onn, ainn, iann, iunn/ionn, uann, uainn, iaunn/iaonn
# （3）鼻音韻尾：
# am, an, ang, im, in, ing, un, ong, iam, ian, iang, iong, uan, uang
UN_BU_TNG_UANN_PIAU = {
    # （1）鼻化韻母
    "iaunn": "niao",
    "uainn": "nuai",
    "uann": "nua",
    "iunn": "niu",
    "ionn": "nio",
    "iann": "nia",
    "ainn": "nai",
    "oonn": "no",
    "onn": "no",
    "enn": "ne",
    "inn": "ni",
    # （2）複合韻母
    # "uai": "uai",
    "iau": "iao",
    "ue": "ue",
    "ui": "ui",
    "ua": "ua",
    "io": "io",
    "iu": "iu",
    "ia": "ia",
    "au": "ao",
    "ai": "ai",
    # （3）鼻音韻尾：
    # am, an, ang, im, in, ing, un, ong, iam, ian, iang, iong, uan, uang
    "uang": "uang",
    "uan": "uan",
    "iong": "iong",
    "iang": "iang",
    "ian": "ian",
    "iam": "iam",
    "ong": "ong",
    "un": "un",
    "ing": "ing",
    "in": "in",
    "im": "im",
    "ang": "ang",
    "an": "an",
    "am": "am",
    # （1）元音及方音
    # "ir": "ir",
    # "ee": "ee",
    # "a": "a",
    # "i": "i",
    # "u": "u",
    # "e": "e",
    # "oo": "oo",
    # "o": "o",  # ㄜ
}

VOWELS = set("aeiou")  # 用於判斷「i/u 後是否接母音」

TLPA_TIAU_HO_TNG_TIAU_MIA = {
    "1": "陰平",
    "2": "陰上",
    "3": "陰去",
    "4": "陰入",
    "5": "陽平",
    "6": "陽上",
    "7": "陽去",
    "8": "陽入",
}
BP_TIAU_MIA_TNG_TIAU_HO = {
    "陰平": "1",
    "陽平": "2",
    "陰上": "3",
    "陽上": "3",
    "陰去": "5",
    "陽去": "6",
    "陰入": "7",
    "陽入": "8",
}


def convert_TLPA_to_BP(TLPA_piau_im: str):
    """
    將一個 【台語音標】（TLPA）轉換成帶【聲調符號】的【閩拼音標】（BP）。如：
    【滾】==> kun3【台語音標】==> gun3【帶調號閩拼音標】。
    若輸入不符合正規表達式 ^([a-z]+)(\\d+)$，則回傳 None。

    轉換步驟（概要）：
    1. 以正規表達式分離拼音（聲母+韻母）與聲調數字。
    2. 以 SIANN_BU_TNG_UANN_PIAU 做最長前綴比對轉換聲母，剩餘為韻母。
    3. 若剩餘韻母整段能在 UN_BU_TNG_UANN_PIAU 找到對應，則以對應值取代。
    4. 處理零聲母且韻母以 i 或 u 起頭的介音情形：
       - i + 母音字母：聲母設為 "y"，刪去韻母首 i。例： 腰 [iao] ==> [yao]，鞅 [iang] ==> [yang]，央 [iong] ==> [yong]
       - i + 非母音字母：聲母設為 "y"，韻母首保留。例： 伊 [i] ==> [yi]，音 [im] ==> [yim]，益 [ik] ==> [yik]
       - u + 母音字母：聲母設為 "w"，刪去韻母首 u。例： 彎 [uan] ==> [wan]，歪 [uai] ==> [wai]，位 [ui] ==> [wi]
       - u + 非母音字母：聲母設為 "w"，韻母首保留。例： 有 [u] ==> [wu]，溫 [un] ==> [wun]，鬱 [ut] ==> [wut]
    5. 以 TLPA_TIAU_HO_PIAU 與 BP_TIAU_HO_PIAU 做聲調名稱與編碼之對應轉換。
    6. 回傳 "<聲母><韻母><聲調>"。

    參數：
    - TLPA_piau_im (str): 要轉換的 TLPA 詞條，如 "tsiann1"、"iao2" 等。

    回傳值：
    - str: 轉換後的 BP 詞條；若輸入格式不符則回傳原字串。
    """
    # 確認傳入之【台語音標】符合格式=聲母+韻母+聲調=英文字母+數字
    m = re.match(r"^([a-z]+)(\d+)$", TLPA_piau_im)
    if not m:
        # 如果不符合「全英文字母+數字」格式，就原樣回傳
        return None, None, None

    # 提取：【無調號標音】（聲母+韻母）和【聲調】
    mo_tiau_piau_im, tiau = m.group(1), m.group(2)

    # 1. 轉聲母：從長到短比對 prefix
    # 特殊處理：韻化聲母 m、ng（後面直接接聲調，不轉換）
    siann = ""
    un = mo_tiau_piau_im

    # 檢查是否為韻化聲母：m 或 ng 後面沒有韻母（整個無調號標音就是 m 或 ng）
    if mo_tiau_piau_im == "m":
        # 韻化聲母 m：毋 [m7] 保持為 m，不轉換成 bbn
        siann = ""
        un = "m"
    elif mo_tiau_piau_im == "ng":
        # 韻化聲母 ng：黃 [ng5] 保持為 ng，不轉換成 ggn
        siann = ""
        un = "ng"
    else:
        # 正常聲母轉換邏輯
        for key in sorted(SIANN_BU_TNG_UANN_PIAU.keys(), key=lambda x: -len(x)):
            if mo_tiau_piau_im.startswith(key):
                siann = SIANN_BU_TNG_UANN_PIAU[key]
                un = mo_tiau_piau_im[len(key) :]
                break

    # 2. 轉韻母：整段比對
    if un in UN_BU_TNG_UANN_PIAU:
        un = UN_BU_TNG_UANN_PIAU[un]

    # 3.【零聲母連i/u】特殊處理
    if siann == "" and un:
        first_lo_ma_ji_bu = un[0]

        if first_lo_ma_ji_bu == "i":
            # i 為【介音】，聲母變更為：[y]，韻母的首羅馬字 [i] 將之刪除。
            # 【例】：腰 [iao] ==> [yao]，鞅 [iang] ==> [yang]，央 [iong] ==> [yong]
            # i 後面是母音：移到聲母 y，刪掉韻母開頭 i（1.2）
            if len(un) >= 2 and un[1] in VOWELS:
                siann = "y"
                un = un[1:]
            else:
                # i 為【元音】韻母，聲母變更為：[y]，韻母維持不變。
                # 【例】：伊 [i] ==> [yi]，音 [im] ==> [yim]，益 [ik] ==> [yik]
                # i 後面不是母音：移到聲母 y，但韻母保留 i（1.1）
                siann = "y"
                # un 保持以 i 起頭，例如 i / in / inn

        elif first_lo_ma_ji_bu == "u":
            # u 為【介音】，聲母變更為：[w]，韻母的首羅馬字 [u] 將之刪除。
            # 【例】：彎 [uan] ==> [wan]，歪 [uai] ==> [wai]，位 [ui] ==> [wi]
            # u 後面是母音：移到聲母 w，刪掉韻母開頭 u（2.2）
            if len(un) >= 2 and un[1] in VOWELS:
                siann = "w"
                un = un[1:]
            else:
                # u 為【元音】韻母，聲母變更為：[w]，韻母維持不變。
                # 【例】：有 [u] ==> [wu]，溫 [un] ==> [wun]，鬱 [ut] ==> [wut]
                # u 後面不是母音：移到聲母 w，但韻母保留 u（2.1）
                siann = "w"
                # un 保持以 u 起頭，例如 u / un / unn

    # 4. 【台語音標】調號轉換成【閩拼音標】調號
    tiau_mia = TLPA_TIAU_HO_TNG_TIAU_MIA.get(tiau, tiau)
    tiau = BP_TIAU_MIA_TNG_TIAU_HO.get(tiau_mia, tiau_mia)
    # return f"{siann}{un}{tiau}"
    return [siann, un, tiau]


def convert_TLPA_to_BP_with_tone_marks(tlpa_piau_im: str) -> str:
    """
    將一個 【台語音標】（TLPA）轉換成帶【聲調符號】的【閩拼音標】（BP）。如：
    【滾】==> kun3【台語音標】==> gun3【帶調號閩拼音標】 ==> gǔn【帶調符閩拚音標】。

    輸入格式：小寫羅馬拚音字母組成的【音節】=【聲母】+【韻母】+【聲調】，例如：
    【滾】==> gun3 = g（聲母） + un（韻母） + 3（聲調）==> 轉換後為 gǔn 。
    若輸入不符合正規表達式 ^([a-z]+)(\\d+)$，則回傳【None】。

    轉換步驟（概要）：
    1. 將【台語音標】轉換成【閩拚音標】；
    2. 【閩拚音標】解構成：【無調號音標】（聲母+韻母）與【調號】：
    3. 若遇【聲母】為【零聲母】，且【韻母】的首字母為 i 或 u 之狀況：
       - i + 母音字母：聲母設為 "y"，刪去韻母首 i。例： 腰 [iao] ==> [yao]，鞅 [iang] ==> [yang]，央 [iong] ==> [yong]
       - i + 非母音字母：聲母設為 "y"，韻母首保留。例： 伊 [i] ==> [yi]，音 [im] ==> [yim]，益 [ik] ==> [yik]
       - u + 母音字母：聲母設為 "w"，刪去韻母首 u。例： 彎 [uan] ==> [wan]，歪 [uai] ==> [wai]，位 [ui] ==> [wi]
       - u + 非母音字母：聲母設為 "w"，韻母首保留。例： 有 [u] ==> [wu]，溫 [un] ==> [wun]，鬱 [ut] ==> [wut]
    4. 在【韻母】中【響度】最高的【元音字母】上，加上【調號】對映的【聲調符號】。
    5. 回傳【帶調符號閩拚音標】字串。

    參數：
    - tlpa_piau_im (str): 待轉換的【台語音標】，如 "tsiann1"、"iao2" 等。

    回傳值：
    - str: 轉換後的【帶調符閩拼音標】；若輸入格式不符則回傳【None】。
    """

    tiau_ho_to_tiau_hu_mapping = {
        0: "\u030A",  # 陰平
        1: "\u0304",  # 陰平
        2: "\u0301",  # 陽平
        3: "\u030C",  # 上声
        5: "\u0300",  # 陰去
        6: "\u0302",  # 陽去
        7: "\u0304",  # 陰入
        8: "\u0301",  # 陽入
    }
    # 聲調符號對應表（帶調號母音 → 對應數字）
    tone_mapping = {
        "å": ("a","0"), "ā": ("a","1"), "á": ("a","2"), "ǎ": ("a","3"), "à": ("a","5"), "â": ("a","6"), "āh": ("a","7"), "áh": ("a","8"), "a̋": ("a","9"),
        "e̊": ("e","0"), "ē": ("e","1"), "é": ("e","2"), "ě": ("e","3"), "è": ("e","5"), "ê": ("e","6"), "ēh": ("e","7"), "éh": ("e","8"), "e̋": ("e","9"),
        "i̊": ("i","0"), "ī": ("i","1"), "í": ("i","2"), "ǐ": ("i","3"), "ì": ("i","5"), "î": ("i","6"), "īh": ("i","7"), "íh": ("i","8"), "i̋": ("i","9"),
        "o̊": ("o","0"), "ō": ("o","1"), "ó": ("o","2"), "ǒ": ("o","3"), "ò": ("o","5"), "ô": ("o","6"), "ōh": ("o","7"), "óh": ("o","8"), "ő": ("o","9"),
        "ů": ("u","0"), "ū": ("u","1"), "ú": ("u","2"), "ǔ": ("u","3"), "ù": ("u","5"), "û": ("u","6"), "ūh": ("u","7"), "úh": ("u","8"), "ű": ("u","9"),
        "m̊": ("m","0"), "m̄": ("m","1"), "ḿ": ("m","2"), "m̌": ("m","3"), "m̀": ("m","5"), "m̂": ("m","6"), "m̄h": ("m","7"), "ḿh": ("m","8"), "m̋": ("m","9"),
        "n̊": ("n","0"), "n̄": ("n","1"), "ń": ("n","2"), "ň": ("n","3"), "ǹ": ("n","5"), "n̂": ("n","6"), "n̄h": ("n","7"), "ńh": ("n","8"), "n̋": ("n","9"),
    }
    # 韻母中元音的響度優先順序（用於決定調符標注位置）
    # 響度從高到低：a > o > e > i/u > m/n
    vowel_priority = {
        'a': 5,
        'o': 4,
        'e': 3,
        'i': 2,
        'u': 2,
        'm': 1,
        'n': 1,
    }

    # 步驟 1: 將【台語音標】轉換成【閩拼音標】
    result = convert_TLPA_to_BP(tlpa_piau_im)
    if result is None or result[0] is None:
        return None

    siann, un, tiau = result

    # 步驟 2: 已完成（在 convert_TLPA_to_BP 中）
    # 步驟 3: 已完成（在 convert_TLPA_to_BP 中）

    # 調號需轉為字串格式以匹配 mapping
    tiau_ho_to_tiau_hu_mapping = {
        "0": "\u030A",  # 陰平（輕聲）
        "1": "\u0304",  # 陰平
        "2": "\u0301",  # 陽平
        "3": "\u030C",  # 上聲
        "5": "\u0300",  # 陰去
        "6": "\u0302",  # 陽去
        "7": "\u0304",  # 陰入
        "8": "\u0301",  # 陽入
    }

    # 步驟 4: 在韻母中響度最高的元音字母上加上聲調符號
    if not un or tiau not in tiau_ho_to_tiau_hu_mapping:
        # 若韻母為空或調號不在對應表中，直接返回
        return f"{siann}{un}{tiau}"

    # 找出韻母中響度最高的元音位置
    max_priority = -1
    target_index = -1

    for i, char in enumerate(un):
        if char in vowel_priority:
            priority = vowel_priority[char]
            if priority > max_priority:
                max_priority = priority
                target_index = i

    # 若找不到元音，直接返回無調符的音標
    if target_index == -1:
        return f"{siann}{un}"

    # 在目標元音上加上調符
    tone_mark = tiau_ho_to_tiau_hu_mapping[tiau]
    target_vowel = un[target_index]

    # 組合：聲母 + 韻母前段 + 帶調符元音 + 韻母後段
    un_with_tone = (
        un[:target_index] +
        target_vowel +
        tone_mark +
        un[target_index + 1:]
    )

    # 步驟 5: 回傳帶調符的閩拼音標
    return f"{siann}{un_with_tone}"