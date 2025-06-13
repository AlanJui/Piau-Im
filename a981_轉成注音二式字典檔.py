#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_to_zu_im_2.py

將 tl_ji_khoo_peh_ue.dict.yaml 之 code（台羅）轉成注音二式，
輸出成 zu_im_2.dict.yaml。
用法：
    python convert_to_zu_im_2.py [輸入檔.yaml] [輸出檔.yaml]
若不帶參數，預設輸入：tl_ji_khoo_peh_ue.dict.yaml，
        輸出：zu_im_2.dict.yaml。
"""

import re
import sys

# 聲母映射表（TL → 注音二式），從長到短
TL_INITIAL_MAP = {
    'tshi': 'ci',
    'tsi':  'zi',
    'tsh':  'c',
    'ts':   'z',
    'ji':   'zzi',
    'j':    'zz',
    'ph':   'p',
    'p':    'b',
    'th':   't',
    't':    'd',
    'kh':   'k',
    'k':    'g',
    'g':    'gg',
    # 'ng':   'ng-',
    'si':   'si',
    's':    's',
    'm':    'm',
    'n':    'n',
    'l':    'l',
    'h':    'h',
}

# 韻母只針對這三種做轉換，其它都保留
TL_FINAL_MAP = {
    'onn': 'oonn',  # ㆧ
    'o':   'or',    # ㄜ
    'io':  'ior',   # 新增 io → ior
}

def convert_tl_to_bopo2(code: str) -> str:
    """
    TL 拼音（小寫英文字母+數字）轉 注音二式。
    只對 TL_INITIAL_MAP 全部轉，TL_FINAL_MAP 裡的三個韻母轉，
    其餘韻母保持 body tone 不變。
    """
    m = re.match(r'^([a-z]+)(\d+)$', code)
    if not m:
        return code
    body, tone = m.group(1), m.group(2)

    # 1. 聲母從長到短比對
    onset, rest = '', body
    for key in sorted(TL_INITIAL_MAP, key=lambda x: -len(x)):
        if body.startswith(key):
            onset = TL_INITIAL_MAP[key]
            rest  = body[len(key):]
            break

    # 2. 韻母只轉 FINAL_MAP 裡定義的
    if rest in TL_FINAL_MAP:
        rest = TL_FINAL_MAP[rest]

    return f"{onset}{rest}{tone}"

def main(in_yaml: str, out_yaml: str):
    with open(in_yaml, 'r', encoding='utf-8') as fin, \
         open(out_yaml, 'w', encoding='utf-8') as fout:

        in_entries = False
        for line in fin:
            # 在「...」之前都原樣輸出
            if not in_entries:
                fout.write(line)
                if line.strip() == '...':
                    in_entries = True
                continue

            # 進入詞條區後
            tpl = line.rstrip('\n')
            if not tpl or tpl.startswith('#'):
                fout.write(line)
                continue

            # 用 tab 分成四欄：text, code, weight, stem
            parts = tpl.split('\t', 3)
            if len(parts) >= 2:
                text = parts[0]
                code = parts[1]
                newcode = convert_tl_to_bopo2(code)
                # 若有 weight/stem，就把剩下的原樣接回來；否則補上空白欄位
                if len(parts) == 4:
                    weight, stem = parts[2], parts[3]
                elif len(parts) == 3:
                    weight, stem = parts[2], ''
                else:
                    weight, stem = '', ''
                fout.write(f"{text}\t{newcode}\t{weight}\t{stem}\n")
            else:
                # 格式不符就直接寫回
                fout.write(line)

    print(f"已輸出新字典：{out_yaml}")

if __name__ == '__main__':
    # 預設檔名
    default_in  = "tl_ji_khoo_peh_ue.dict.yaml"
    default_out = "zu_im_2.dict.yaml"
    if len(sys.argv) == 3:
        inp, outp = sys.argv[1], sys.argv[2]
    else:
        inp, outp = default_in, default_out
        print(f"未指定檔名，使用預設：\n  輸入 → {inp}\n  輸出 → {outp}")
    main(inp, outp)
