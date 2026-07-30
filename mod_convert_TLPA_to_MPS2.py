"""
convert_TLPA_to_MPS2.py

將【台語音標（TLPA+）】轉換成【台語注音二式（MPS2）】。
用法：
    python convert_TLPA_to_MPS2.py tl_ji_khoo_peh_ue.txt mps2_ji_khoo.txt
"""

import re
import sys

# 聲母轉換對照表（【索引】字串排序，需由長到短）
SIANN_BU_TNG_UANN_PIAU = {
    "tsh": "c",
    "ts": "z",
    # 二字母
    "ph": "p",  # ㄆ → p (雙唇音/清音：塞音/送氣)
    "th": "t",  # ㄊ → t (齒齦音/清音：塞音/送氣)
    "kh": "k",  # ㄎ → k（軟顎音/清音：塞音/送氣）
    "ng": "ng",  # ㄫ → ng（軟顎音/濁音：鼻音）
    # 一字母
    # 雙唇音
    "p": "b",  # ㄅ → b（雙唇音/清音：塞音不送氣）
    "b": "bb",  # ㆠ → bb（雙唇音/濁音：塞音不送氣）
    "m": "m",  # ㄇ → m（雙唇音/濁音：鼻音）
    # ------------------------------
    # 齒齦音
    "t": "d",  # ㄉ → d（齒齦音/清音：塞音/不送氣）
    "n": "n",  # ㄋ → n（齒齦音/濁音：鼻音）
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
CI_IM_TNG_UANN_PIAU = {
    "zzi": "jji",  # ㆢ：ji → jj+i
    "zi": "ji",  # ㄐ：z+i → j+i
    "ci": "chi",  # ㄑ：c+i → ch+i
    "si": "shi",  # ㄒ：s+i → sh+i
}

# 韻母轉換對照表
# 【準則】以 Documents/_聲韻調對照表.xlsx【韻】韻母對照表為標準；
#         此處僅列出「台語音標 != 台語注音二式」之韻母（需轉換者），
#         未列出者一律維持原樣（identity）。採整段比對（非前綴比對）。
UN_BU_TNG_UANN_PIAU = {
    # ㆤ（漳腔）→ e
    "ee": "e",
    "eeh": "eh",
    "ei": "e",
    # ㄜ（o → or）系列
    "o": "or",
    "oh": "orh",
    "io": "ior",
    "ioh": "iorh",
    # ㆦ（oo）鼻化／入聲系列
    "om": "oom",
    "op": "oop",
    "onn": "oonn",
    "ohnn": "onnh",
    "ionn": "ioonn",
    # ㄧㆤ（ik → iek）、ㄧㆦㆻ（iok → iook）
    "ik": "iek",
    "iok": "iook",
    # 鼻化＋入聲（台語音標 Vhnn → 注音二式 Vnnh，h 移至韻尾）
    "ahnn": "annh",
    "ehnn": "ennh",
    "ihnn": "innh",
    "uaihnn": "uainnh",
    "iauhnn": "iaunnh",
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

    # 提取：【無調號標音】（聲母+韻母）和【聲調】
    mo_tiau_piau_im, tiau = m.group(1), m.group(2)

    # 1. 轉聲母：從長到短比對 prefix
    siann = ""
    un = mo_tiau_piau_im
    for key in sorted(SIANN_BU_TNG_UANN_PIAU.keys(), key=lambda x: -len(x)):
        if mo_tiau_piau_im.startswith(key):
            siann = SIANN_BU_TNG_UANN_PIAU[key]
            un = mo_tiau_piau_im[len(key) :]
            break

    # 2. 轉韻母：整段比對
    if un in UN_BU_TNG_UANN_PIAU:
        un = UN_BU_TNG_UANN_PIAU[un]
    # else:
    #     # 若末尾是「o」卻不在 FINAL_MAP，做一次 o→or
    #     if rest.endswith("o"):
    #         rest = rest[:-1] + "or"

    # 3. 處理【齒音連i】的特殊狀況：【齒音聲母】+ i（【韻母】首拼音字母）
    if siann in ("z", "c", "s", "zz") and un.startswith("i"):
        ci_im_lian_i = f"{siann}i"
        if ci_im_lian_i in CI_IM_TNG_UANN_PIAU:
            ci_im_lian_i_tng_uann = CI_IM_TNG_UANN_PIAU[ci_im_lian_i]
            siann = ci_im_lian_i_tng_uann[:-1]  # 去掉最後的 i

    return f"{siann}{un}{tiau}"


# ---------------------------------------------------------------------
# MPS2（台語注音二式）→ TLPA 反轉換
# ---------------------------------------------------------------------
# 聲母反轉換（由長到短比對）；齒音連 i（jj/j/ch/sh）另於正文前先還原。
MPS2_TO_TLPA_SIANN = {
    "bb": "b",
    "gg": "g",
    "zz": "j",
    "ng": "ng",
    "b": "p",
    "d": "t",
    "g": "k",
    "p": "ph",
    "t": "th",
    "k": "kh",
    "c": "c",
    "z": "z",
    "m": "m",
    "n": "n",
    "l": "l",
    "s": "s",
    "h": "h",
}

# 齒音連 i：MPS2 → TLPA（還原 CI_IM_TNG_UANN_PIAU）
MPS2_CI_IM_REV = {
    "jji": "zzi",
    "ji": "zi",
    "chi": "ci",
    "shi": "si",
}

# 韻母反轉換：略過 ee/ei→e、eeh→eh 等歧義項（還原時維持 e/eh）
MPS2_TO_TLPA_UN = {
    mps2: tlpa
    for tlpa, mps2 in UN_BU_TNG_UANN_PIAU.items()
    if not (tlpa in {"ee", "ei"} and mps2 == "e") and not (tlpa == "eeh" and mps2 == "eh")
}


def convert_MPS2_to_TLPA(MPS2_piau_im: str) -> str:
    """
    將一個【注音二式/MPS2】（如 'ziann1'）轉回【台語音標/TLPA】（'tsiann1' 之前的 TLPA 形：'ziann1'）。
    保留後面的數字（聲調）。
    """
    m = re.match(r"^([a-z]+)(\d+)$", MPS2_piau_im)
    if not m:
        return MPS2_piau_im

    body, tiau = m.group(1), m.group(2)

    # 1. 先還原【齒音連 i】（須由長到短）
    for mps2_ci, tlpa_ci in sorted(MPS2_CI_IM_REV.items(), key=lambda x: -len(x[0])):
        if body.startswith(mps2_ci):
            body = tlpa_ci + body[len(mps2_ci) :]
            break

    # 2. 反轉聲母
    siann = ""
    un = body
    for key in sorted(MPS2_TO_TLPA_SIANN.keys(), key=lambda x: -len(x)):
        if body.startswith(key):
            siann = MPS2_TO_TLPA_SIANN[key]
            un = body[len(key) :]
            break

    # 3. 反轉韻母（整段比對，長鍵優先）
    for key in sorted(MPS2_TO_TLPA_UN.keys(), key=lambda x: -len(x)):
        if un == key:
            un = MPS2_TO_TLPA_UN[key]
            break

    return f"{siann}{un}{tiau}"


def main(infile: str, outfile: str):
    with open(infile, "r", encoding="utf-8") as fin:
        lines = fin.readlines()

    in_entries = False
    out_lines = []
    for line in lines:
        # 找到「...」之後即進入詞條區
        if not in_entries:
            out_lines.append(line)
            if line.strip() == "...":
                in_entries = True
            continue

        # 略過註解行或空白行
        if not line.strip() or line.startswith("#"):
            out_lines.append(line)
            continue

        # 轉換【第二欄】中的【台語音標】
        parts = line.rstrip("\n").split("\t")  # 假設詞條以「欄位1\t欄位2\t...」格式，至少要有兩欄
        if len(parts) >= 2:
            # 轉換第二欄中之【台語音標】，並將結果覆蓋回去
            parts[1] = convert_TLPA_to_MPS2(parts[1])
            # 將轉換結果放入【轉換結果】陣列底端
            out_lines.append("\t".join(parts) + "\n")
        else:
            out_lines.append(line)

    # 將轉換結果寫入輸出檔
    with open(outfile, "w", encoding="utf-8") as fout:
        fout.writelines(out_lines)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法：python convert_TLPA_to_MPS2.py <輸入檔> <輸出檔>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
