#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import re

import requests

API = "https://www.moedict.tw/a/{}.json"

INITIALS = [
    "ph", "th", "kh", "chh",
    "p", "t", "k", "m", "n", "ng", "l", "s", "h", "ch", "j", "b", "g"
]

TONE_PATTERN = re.compile(r"([0-9])$")


def split_tl(tl: str):
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
    """查萌典華台對照 API"""
    url = API.format(keyword)
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    return resp.json()


def main():
    parser = argparse.ArgumentParser(description="台語字典 CLI（萌典華台對照版）")
    parser.add_argument("word", help="要查的詞，例如：白")
    args = parser.parse_args()

    try:
        data = query(args.word)
    except Exception as e:
        print(f"查詢失敗：{e}")
        return

    title = data.get("title")
    heteronyms = data.get("heteronyms", [])

    print(f"詞目：{title}")

    if heteronyms:
        h = heteronyms[0]
        tl = h.get("tl")
        poj = h.get("poj")
        bopomofo = h.get("bopomofo")

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
