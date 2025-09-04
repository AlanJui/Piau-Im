#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_TLPA_to_BP.py

將【台語音標（TLPA+）】轉換成【閩拼方案（BP）】。
用法：
    python convert_TLPA_to_BP.py <輸入檔> <輸出檔>
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
    "c": "c",    # ㄘ → c (台羅：tshi)
    "z": "z",    # ㄗ → z (台羅：tsi)
    "ph": "p",   # ㄆ → p (台羅：ph)
    "th": "t",   # ㄊ → t (台羅：th)
    "kh": "k",   # ㄎ → k (台羅：kh)
    "ji": "zzi", # ㆢ → zzi (台羅：ji)
    "si": "si",  # ㄒ → si
    # 一字母
    "b": "bb",   # ㆠ → bb
    "p": "b",    # ㄅ → b
    "m": "m",    # ㄇ → m
    "t": "d",    # ㄉ → d
    "n": "n",    # ㄋ → n
    "l": "l",    # ㄌ → l
    "k": "g",    # ㄍ → g
    "g": "gg",   # ㆣ → gg
    "h": "h",    # ㄏ → h
    "j": "zz",   # ㆡ → zz
    "s": "s",    # ㄙ → s
}

# 韻母（襯聲）映射表
# 註：此處「o」特例改為「or」
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
    # 如果字典裡有「-ng」「-ing」「-m」等，也可加進來
    "-ng": "-ng",
    "ing": "ing",
    "m": "m",
}

VOWELS = set("aeiou")  # 用於判斷「i/u 後是否接母音」

def convert_TLPA_to_BP(tai_gi_im_piau: str) -> str:
    """
    將一個【台語音標/TLPA】（如 'tsiann1'）轉成【閩拼/BP】（例如 'ziann1'）。
    保留後面的數字（聲調）。
    """
    m = re.match(r"^([a-z]+)(\d+)$", tai_gi_im_piau)
    if not m:
        # 若不符合「英文字母+數字」格式，就原樣回傳
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

    # 2) 先做韻母映射（含 o→or 特例）
    if rest in FINAL_MAP:
        rest = FINAL_MAP[rest]
    else:
        # 例如 ...o 結尾要變 or
        if rest.endswith("o"):
            rest = rest[:-1] + "or"

    # 3)【零聲母 + i/u】規則（依您新指示精修）：
    #   - 若 onset 為空字串，且 rest 有內容，才檢查
    if onset == "" and rest:
        # 取韻母第一字母
        first = rest[0]

        # 3.1 處理 i
        if first == "i":
            # 3.1.1 若 i 後接母音（a/e/i/o/u），把 i 改到聲母（= y），並消去韻母裡的 i
            if len(rest) >= 2 and rest[1] in VOWELS:
                onset = "y"
                rest = rest[1:]  # 去掉開頭的 i
            else:
                # 3.1.2 若 i 後沒有接母音（如 i, in, inn ...），在韻母前補 y
                if not rest.startswith("yi"):
                    rest = "y" + rest

        # 3.2 處理 u
        elif first == "u":
            # 3.2.1 若 u 後接母音（a/e/i/o/u），把 u 改到聲母（= w），並消去韻母裡的 u
            if len(rest) >= 2 and rest[1] in VOWELS:
                onset = "w"
                rest = rest[1:]  # 去掉開頭的 u
            else:
                # 3.2.2 若 u 後沒有接母音（如 u, un, unn ...），在韻母前補 w
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
            parts[1] = convert_TLPA_to_BP(parts[1])
            out_lines.append("\t".join(parts) + "\n")
        else:
            out_lines.append(line)

    with open(outfile, "w", encoding="utf-8") as fout:
        fout.writelines(out_lines)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法：python convert_TLPA_to_BP.py <輸入檔> <輸出檔>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
