import unicodedata

FULLWIDTH_PUNCTUATIONS = set("，。、；：？！﹁﹂﹃﹄（）［］｛｝「」『』《》〈〉【】～｜‧＂＇…—")

def kam_si_piau_tian(char):
    return char in FULLWIDTH_PUNCTUATIONS

# 用途：檢查是否為漢字
def is_han_ji(char):
    return 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(char, '')

def kam_si_cjk_han_ji(char):
    return (
        '\u4E00' <= char <= '\u9FFF' or
        '\u3400' <= char <= '\u4DBF' or
        '\U00020000' <= char <= '\U0002A6DF' or
        '\U0002A700' <= char <= '\U0002B73F' or
        '\U0002B740' <= char <= '\U0002B81F' or
        '\U0002B820' <= char <= '\U0002CEAF' or
        '\U0002CEB0' <= char <= '\U0002EBEF'
    )

def ut01():
    # 測試
    chars = ['？', '！', '，', '。', '：', '；', '（', '）', '【', '】', '『', '』', '東', 'A', '?']

    for char in chars:
        print(f"{char}: {kam_si_piau_tian(char)}")

def ut02():
    # 此測試結果會失敗
    han_ji_list = ['東', 'A', '?', '？']
    for han_ji in han_ji_list:
        print(f"han_ji: {han_ji} {is_han_ji(han_ji)}")

def ut03():
    # 此測試結果會失敗
    han_ji_list = ['東', 'A', '?', '？']
    for han_ji in han_ji_list:
        print(f"han_ji: {han_ji} {kam_si_cjk_han_ji(han_ji)}")

if __name__ == "__main__":
    # ut01()
    # ut02()
    ut03()








