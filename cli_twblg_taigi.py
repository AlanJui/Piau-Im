#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import re

import requests

# 台文華文線頂典 API（穩定、無封鎖）
API = "https://twblg.dict.edu.tw/holodict_new/"

INITIALS = [
    "ph", "th", "kh", "chh",
    "p", "t", "k", "m", "n", "ng", "l", "s", "h", "ch", "j", "b", "g"
]

TONE_PATTERN = re.compile(r"([0-9])$")


def split_tl(tl: str):
    """拆分台羅：聲母、韻母、調號"""
    tl = tl.lower().strip()
    tone_match = TONE_PATTERN.search(tl)
    tone = tone_match.group(1) if tone_match else ""
    syllable = tl[:-1] if tone else tl

    initial = ""
    for ini in sorted(INITIALS, key=len, reverse=True):
        if syllable.startswith(ini):
            initial = ini
            break

    final = syllable[len(initial):]
    return initial, final, tone


def query(keyword: str):
    """查詢台文華文線頂典 API"""

    params = {
        "op": "search",
        "begin": 1,
        "page": 1,
        "result_num": 20,
        "search": keyword,
    }

    resp = requests.get(API, params=params, timeout=10)
    resp.raise_for_status()

    data = resp.json()

    # API 回傳格式：
    # {
    #   "result": [
    #       {
    #           "title": "白",
    #           "poj": "pe̍h",
    #           "tl": "peh8",
    #           "bopomofo": "ㄅㄧㄚㆷ˪",
    #           ...
    #       }
    #   ]
    # }
    return data.get("result", [])


def main():
    parser = argparse.ArgumentParser(description="台語字典 CLI（台文華文線頂典版）")
    parser.add_argument("word", help="要查的詞，例如：白")
    args = parser.parse_args()

    try:
        entries = query(args.word)
    except Exception as e:
        print(f"查詢失敗：{e}")
        return

    if not entries:
        print("查無資料")
        return

    entry = entries[0]

    title = entry.get("title")
    tl = entry.get("tl")
    poj = entry.get("poj")
    bopomofo = entry.get("bopomofo")

    print(f"詞目：{title}")
    print(f"台羅：{tl}")
    print(f"白話字：{poj}")
    print(f"方音符號：{bopomofo}")

    if tl:
        ini, fin, tone = split_tl(tl)
        print("\n【台羅拆音】")
        print(f"聲母：{ini}")
        print(f"韻母：{fin}")
        print(f"調號：{tone}")


if __name__ == "__main__":
    main()
