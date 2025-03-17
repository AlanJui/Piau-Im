import unicodedata


def unicode帶調拼音字母():
    # 帶調拼音字母 á ，有兩種方式表示：（1）單一字元（2）基本字母 + 聲調符號
    print("\u00E1")     # á (U+00E1)
    print("a\u0301")    # á (U+0061 + U+0301)

def unicode無法帶調拼音字母():
    # 陽入聲調羅馬字母，只能使用：基本字母 + 聲調符號
    a2_2_char = "a\u030D"    # 	a̍ (U+0061 + U+030D)
    print(a2_2_char)

def list使用():
    # tiau_hu_list = ["\u0301", "\u0300", "\u0302", "\u0304", "\u030D"]
    tiau_hu_list = (
        "\u0301", # 2：陰上
        "\u0300", # 3：陰去
        "\u0302", # 5：陽平
        "\u0304", # 7：陽去
        "\u030D", # 8：陽入
    )
    print("--------------------------------------------------------------------")
    idx = 0
    un_a = []
    for tiau in tiau_hu_list:
        un_a.append(f"a{tiau}")
        print(f"{idx+1}. a{tiau}")
        idx += 1
    print("--------------------------------------------------------------------")
    # 使用 for ... in [List] 來取得元素：不透過索引值，直接取得元素
    idx = 0
    for un in un_a:
        print(f"{idx+1}. {un}")
        idx += 1
    print("--------------------------------------------------------------------")
    # 使用 for ... in range(len([List])) 來取得元素：透過索引值取得元素
    idx = 0
    for i in range(len(un_a)):
        print(f"{idx+1}. {un_a[i]}")
        idx += 1
    print("--------------------------------------------------------------------")


def process():
    x1 = unicodedata.normalize("NFD", "a\u030D")  # 先正規化，拆解聲調符號
    x2 = unicodedata.normalize("NFD", "a\u030D")  # 先正規化，拆解聲調符號
    print(f"x1 = {x1}")
    print(f"x2 = {x2}")
    print("--------------------------------------------------------------------")

    y1 = unicodedata.normalize("NFC", "a\u030D")  # 先正規化，拆解聲調符號
    y2 = unicodedata.normalize("NFC", "a\u030D")  # 先正規化，拆解聲調符號
    print(f"y1 = {y1}")
    print(f"y2 = {y2}")
    print("--------------------------------------------------------------------")

if __name__ == "__main__":
    process()
    print("--------------------------------------------------------------------")
    print("Done.")