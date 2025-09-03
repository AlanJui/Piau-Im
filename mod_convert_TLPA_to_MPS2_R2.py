"""
convert_TLPA_to_MPS2.py

將【台語音標（TLPA+）】轉換成【台語注音二式（MPS2）】。
用法：
    python convert_TLPA_to_MPS2.py tl_ji_khoo_peh_ue.dict.yaml output.dict.yaml
"""

import re
import sys

# 聲母（初聲）映射表，注意要從長到短比對 prefix
INITIAL_MAP = {
    # 四字母（如果有更長的 key，也放最前面）
    "ci": "ci",  # ㄑ → ci (台羅：tshi)
    # 三字母
    "zi": "zi",  # ㄐ → zi (台羅：tsi)
    # 二字母
    "c": "c",  # ㄘ → c (台羅：tshi)
    "z": "z",  # ㄗ → z (台羅：tsi)
    "ph": "p",  # ㄆ → p (台羅：ph)
    "th": "t",  # ㄊ → t (台羅：th)
    "kh": "k",  # ㄎ → k (台羅：kh)
    "ji": "zzi",  # ㆢ → zzi (台羅：ji)
    "si": "si",  # ㄒ → si
    # 一字母
    "b": "bb",  # ㆠ → bb
    "p": "b",  # ㄅ → b
    "m": "m",  # ㄇ → m
    "t": "d",  # ㄉ → d
    "n": "n",  # ㄋ → n
    "l": "l",  # ㄌ → l
    "k": "g",  # ㄍ → g
    "g": "gg",  # ㆣ → gg
    "h": "h",  # ㄏ → h
    "j": "zz",  # ㆡ → zz
    "s": "s",  # ㄙ → s
}

# 韻母（襯聲）映射表，台羅→注音二式（多數相同，唯「o」→「or」需要特別處理）
FINAL_MAP = {
    "i": "i",
    "inn": "inn",
    "u": "u",
    "unn": "unn",
    "a": "a",
    "ann": "ann",
    "oo": "oo",
    "oonn": "oonn",
    "o": "or",  # ㄜ
    "e": "e",
    "enn": "enn",
    "ai": "ai",
    "ainn": "ainn",
    "au": "au",
    "aunn": "aunn",
    "an": "an",
    "en": "en",
    "ang": "ang",
    "ir": "ir",
    "am": "am",
    "om": "om",
    "ong": "ong",
    # 如果你的字典裡有「-ng」或「-ing」「-m」等，也可加進來：
    "-ng": "-ng",
    "ing": "ing",
    "m": "m",
}


def convert_TLPA_to_MPS2(tai_gi_im_piau: str) -> str:
    """
    將一個【台語音標/TLPA】（如 'tsiann1'）轉成【注音二式/MPS2】（'ziann1'）。
    保留後面的數字（聲調）。
    """
    m = re.match(r"^([a-z]+)(\d+)$", tai_gi_im_piau)
    if not m:
        # 如果不符合「全英文字母+數字」格式，就原樣回傳
        return tai_gi_im_piau

    body, tone = m.group(1), m.group(2)

    # 1) 轉聲母：從長到短比對 prefix
    onset = ""
    rest = body
    for key in sorted(INITIAL_MAP.keys(), key=lambda x: -len(x)):
        if body.startswith(key):
            onset = INITIAL_MAP[key]
            rest = body[len(key):]
            break

    # 2) 轉韻母：整段比對（含特例 o→or）
    if rest in FINAL_MAP:
        rest = FINAL_MAP[rest]
    else:
        if rest.endswith("o"):
            rest = rest[:-1] + "or"

    # 2b)【零聲母補 y/w】規則：
    #    無聲母（onset == ""）且韻母以 i / u 起頭，需補 yi / wu
    if onset == "" and rest:
        first = rest[0]
        if first == "i":
            # 避免重複補：若已經是 yi… 就不再加
            if not rest.startswith("yi"):
                rest = "y" + rest
        elif first == "u":
            if not rest.startswith("wu"):
                rest = "w" + rest

    return f"{onset}{rest}{tone}"

def main(infile: str, outfile: str):
    with open(infile, "r", encoding="utf-8") as fin:
        lines = fin.readlines()

    out_lines = []
    in_entries = False
    for line in lines:
        # 找到「...」之後即進入詞條區
        if not in_entries:
            out_lines.append(line)
            if line.strip() == "...":
                in_entries = True
            continue

        # 在詞條區，跳過空行或註解
        if not line.strip() or line.startswith("#"):
            out_lines.append(line)
            continue

        # 假設詞條以「欄位1\t欄位2\t...」格式，至少要有兩欄
        parts = line.rstrip("\n").split("\t")
        if len(parts) >= 2:
            parts[1] = convert_TLPA_to_MPS2(parts[1])
            out_lines.append("\t".join(parts) + "\n")
        else:
            out_lines.append(line)

    with open(outfile, "w", encoding="utf-8") as fout:
        fout.writelines(out_lines)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法：python convert_TL_to_MPS2.py <輸入檔> <輸出檔>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
