import re

valid_un_bu_endings = [
    'un', 'ian', 'im', 'ui', 'ee', 'an', 'ong', 'uai', 'ing', 'uan',
    'oo', 'iau', 'ei', 'iong', 'o', 'ai', 'in', 'iang', 'am', 'ua',
    'ang', 'iam', 'au', 'ia', 'ue', 'ann', 'u', 'a', 'i', 'iu', 'enn',
    'uinn', 'io', 'inn', 'ionn', 'iannh', 'uann', 'ng', 'e', 'ainn',
    'onn', 'm', 'uang', 'uainn', 'uenn', 'iaunn', 'om', 'aunn', 'onn',
    'iunn'
]


def within_tiau_ho(ping_im):
    """
    判斷注音符號中是否含有「聲調」.

    若最後一個字元不是數值，表示使用者可能引用「略去聲調」不寫規則.
    """
    last_char = ping_im[-1]
    return last_char.isdigit()

def split_chu_im(ping_im):
    """
    此方法用於將「台羅拼音」(ping_im) 分解成： 聲母、韻母和調號.

    1. 使用正則表達式（regular expression）匹配聲母。這裡，聲母是由特定字
    符組成的，例如 "b", "tsh", "ts" 等。

    2. 韻母是在聲母之後、調號之前的部分。為了找到韻母，我們首先計算聲母的
    長度（len(siann_bu)），然後從音節的開頭去掉聲母部分，並在音節的末尾
    去掉調號部分。

    3. 調號是音節最後一個字符。

    4. 調號可省略規則：
       當「韻母」為「舒聲韻」時，若「聲調」未標示，則代表「第一聲」；
       當「韻母」為「入聲韻」時，若「聲調」未標示，則代表「第四聲」。
    """
    result = []

    # 正規表達式，用於表達所有可能的聲母。
    siann_pattern = re.compile(r"(b|tsh|ts|g|h|j|kh|k|l|m|ng|n|ph|p|s|th|t|q)?")
    # 透過 match 方法，找到「注音」之中的「聲母」。然後再利用 group
    # 方法，將注音群分「聲母」與「韻母」。
    siann_match = siann_pattern.match(ping_im)

    if siann_match:
        siann_bu = siann_match.group()
    else:
        siann_bu = ""

    # 依據「注音符號」中是否有含「聲調」，決定取得韻母與調號的方式。
    if within_tiau_ho(ping_im):
        # 若注音符號最後一個字元為「數值」，表「聲調」。即
        un_bu = ping_im[len(siann_bu): -1]
        tiau = ping_im[-1]
    else:
        un_bu = ping_im[len(siann_bu):]
        if un_bu in valid_un_bu_endings:
            tiau = '1'
        else:
            tiau = '4'

    result += [siann_bu]
    result += [un_bu]
    result += [tiau]
    return result
