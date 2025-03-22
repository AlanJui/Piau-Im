import unicodedata


def create_tiau_hu_mapping_horizontal():
    uan_im_ji_bu = ("a", "A", "e", "E", "i", "I", "o", "O", "u", "U", "m", "M", "n", "N")

    tiau_fu_tng_tiau_ho_mapping_dict = {
        "\u0301": "2",  # ˊ
        "\u0300": "3",  # ˋ
        "\u0302": "5",  # ˆ
        "\u030C": "6",  # ˇ
        "\u0304": "7",  # ˉ
        "\u0304 ": "8", # 陽入特殊
        "\u030B": "9",  # 雙上標
    }

    print("tiau_hu_mapping = {")

    for uan_im in uan_im_ji_bu:
        print(f"    # {uan_im}")
        row_items = []
        for tiau_fu, tiau_ho in tiau_fu_tng_tiau_ho_mapping_dict.items():
            uan_im_tiau = unicodedata.normalize("NFC", f"{uan_im}{tiau_fu}")
            uan_im_tiau = uan_im_tiau.replace(' ', '')
            row_items.append(f'"{uan_im_tiau}": ("{uan_im}", "{tiau_ho}")')
        print("    " + ", ".join(row_items) + ",")

    print("}")


def create_tiau_hu_mapping_vertical():
    # 基本元音與韻化輔音
    uan_im_ji_bu = ("a", "A", "e", "E", "i", "I", "o", "O", "u", "U", "m", "M", "n", "N")

    # 聲調符號與數字的對應表
    tiau_fu_tng_tiau_ho_mapping_dict = {
        "\u0301": "2",  # ˊ
        "\u0300": "3",  # ˋ
        "\u0302": "5",  # ˆ
        "\u030C": "6",  # ˇ
        "\u0304": "7",  # ˉ
        "\u0304 ": "8", # 特殊陽入調
        "\u030B": "9",  # 雙上標聲調
    }

    # 開始建立 dict
    print("tiau_hu_mapping = {")

    for uan_im in uan_im_ji_bu:
        for tiau_fu, tiau_ho in tiau_fu_tng_tiau_ho_mapping_dict.items():
            # 組合並轉換為單一字元（NFC）
            uan_im_tiau = unicodedata.normalize("NFC", f"{uan_im}{tiau_fu}")
            # 調整沒有成功組合的特殊情況（如陽入符號後面多空白的處理）
            uan_im_tiau = uan_im_tiau.replace(' ', '')
            # 輸出字典項
            print(f'    "{uan_im_tiau}": ("{uan_im}", "{tiau_ho}"),')

    print("}")

def generate_ping_im_regex():
    # 定義基本拼音字母（韻母包含 m, n）
    base_letters = "aeioumn"

    # 定義聲調符號
    accent_marks = "\u0300\u0301\u0302\u0304\u0306\u030C\u030D"

    # 生成所有拼音字母的聲調變體
    pinyin_variants = "".join(f"{ch}{accent}" for ch in base_letters for accent in accent_marks)

    # 動態生成正規表示式
    regex_pattern = rf"[{pinyin_variants}]+$"

    # 顯示結果
    print(regex_pattern)



def ut01():
    uan_im_ji_bu = ("a", "A", "e", "E", "i", "I", "o", "O", "u", "U", "m", "M", "n", "N")

    tiau_fu_tng_tiau_ho_mapping_dict = {
        "": "1",        # 陰平調
        "\u0301": "2",  # 陰上調
        "\u0300": "3",  # 陰去調
        "\u0302": "5",  # 陽平調
        "\u030C": "6",  # 陽上調
        "\u0304": "7",  # 陽去調
        "\u030D": "8",  # 陽入調
        "\u030B": "9",  # 輕聲調
    }
    # tiau_fu_tng_tiau_ho_mapping_dict = {
    #     "\u030D": "8", # 陽入特殊
    #     "\u0301": "2",  # ˊ
    #     "\u030C": "6",  # ˇ
    #     "\u0302": "5",  # ˆ
    #     "\u0304": "7",  # ˉ
    #     "\u0300": "3",  # ˋ
    # }

    print("tiau_hu_mapping = {")

    for uan_im in uan_im_ji_bu:
        # print(f"    # {uan_im}")
        row_items = []
        for tiau_fu, tiau_ho in tiau_fu_tng_tiau_ho_mapping_dict.items():
            uan_im_tiau = unicodedata.normalize("NFC", f"{uan_im}{tiau_fu}")
            uan_im_tiau = uan_im_tiau.replace(' ', '')
            row_items.append(f'"{uan_im_tiau}": ("{uan_im}", "{tiau_ho}")')
        print("    " + ", ".join(row_items) + ",")

    print("}")

if __name__ == "__main__":
    generate_ping_im_regex()