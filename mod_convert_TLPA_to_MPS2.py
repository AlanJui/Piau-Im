"""
convert_TLPA_to_MPS2.py

將【台語音標（TLPA+）】轉換成【台語注音二式（MPS2）】。
用法：
    python convert_TLPA_to_MPS2.py tl_ji_khoo_peh_ue.dict.yaml output.dict.yaml
"""

import re
import sys

# 聲母（初聲）映射表，注意要從長到短比對 prefix
SIANN_BU_MAP = {
    "tsh": "c",
    "ts": "z",
    # 二字母
    "ph": "p",  # ㄆ → p (台羅：ph)
    "th": "t",  # ㄊ → t (台羅：th)
    "kh": "k",  # ㄎ → k (台羅：kh)
    "ng": "ng",  # ㆣ → ng
    # 一字母
    "p": "b",  # ㄅ → b
    "b": "bb",  # ㆠ → bb
    "m": "m",  # ㄇ → m
    "t": "d",  # ㄉ → d
    "n": "n",  # ㄋ → n
    "l": "l",  # ㄌ → l
    "k": "g",  # ㄍ → g
    "g": "gg",  # ㆣ → gg
    "h": "h",  # ㄏ → h
    "z": "z",  # ㄗ → z (台羅：tsi)
    "j": "zz",  # ㆡ → zz
    "c": "c",  # ㄘ → c (台羅：tshi)
    "s": "s",  # ㄙ → s
}

# 齒音（TLPA: 舌尖前音/TL: 舌齒音）+ i 對照：
CI_IM_MAP = {
    # "j": "zz",   # ㆡ：j -> zz
    "zzi": "jji",  # ㆢ：ji → jj+i
}

# 韻母（襯聲）映射表，台羅→注音二式（多數相同，唯「o」→「or」需要特別處理）
FINAL_MAP = {
    "oonn": "oonn",
    # "ainn": "ainn",
    # "aunn": "aunn",
    # "ang": "ang",
    # "ann": "ann",
    # "inn": "inn",
    # "unn": "unn",
    # "enn": "enn",
    # "ong": "ong",
    # "ing": "ing",
    "oo": "oo",
    "ik": "iek",
    # "ai": "ai",
    # "au": "au",
    # "an": "an",
    # "en": "en",
    # "ir": "ir",
    # "am": "am",
    # "om": "om",
    # "a": "a",
    # "i": "i",
    # "u": "u",
    # "e": "e",
    "o": "or",  # ㄜ
}


def convert_TLPA_to_MPS2(TLPA_piau_im: str) -> str:
    """
    將一個【台語音標/TLPA】（如 'tsiann1'）轉成【注音二式/MPS2】（'ziann1'）。
    保留後面的數字（聲調）。
    """
    # 確認傳入之【台語音標】符合格式=聲母+韻母+聲調=英文字母+數字
    m = re.match(r"^([a-z]+)(\d+)$", TLPA_piau_im)
    if not m:
        # 如果不符合「全英文字母+數字」格式，就原樣回傳
        return TLPA_piau_im

    # 提取：聲母、韻母和聲調
    body, tone = m.group(1), m.group(2)

    # 1. 轉聲母：從長到短比對 prefix
    onset = ""
    rest = body
    siann = ""
    for key in sorted(SIANN_BU_MAP.keys(), key=lambda x: -len(x)):
        if body.startswith(key):
            onset = SIANN_BU_MAP[key]
            siann = SIANN_BU_MAP[key]
            rest = body[len(key) :]
            break

    # 2. 轉韻母：整段比對
    if rest in FINAL_MAP:
        rest = FINAL_MAP[rest]
    # else:
    #     # 若末尾是「o」卻不在 FINAL_MAP，做一次 o→or
    #     if rest.endswith("o"):
    #         rest = rest[:-1] + "or"

    # 3. 處理【齒音+ i】的特殊規則
    if siann in ("z", "c", "s", "zz") and rest.startswith("i"):
        ci_im_ga_i = f"{siann}i"
        if ci_im_ga_i in CI_IM_MAP:
            str = CI_IM_MAP[ci_im_ga_i]
            onset = str[:-1]  # 去掉最後的 i

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
        print("用法：python convert_TLPA_to_MPS2.py <輸入檔> <輸出檔>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
