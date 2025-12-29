#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import os
import re

import requests

# ============================
#  教育部 API（目前多半會被擋，先留殼）
# ============================

EDU_API = "https://sutian.moe.edu.tw/api/v1/entries"


def query_edu(keyword: str):
    """嘗試查教育部台語辭典；若被擋或失敗，回傳 None。"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
            "Referer": "https://sutian.moe.edu.tw/",
        }
        resp = requests.get(
            EDU_API,
            params={"search": keyword},
            headers=headers,
            timeout=5,
        )

        # 被 Cloudflare 轉成 HTML 等 → 視為失敗
        if "text/html" in resp.headers.get("Content-Type", ""):
            return None

        data = resp.json()
        entries = data.get("entries", [])
        if not entries:
            return None

        entry = entries[0]
        heteronyms = entry.get("heteronyms", [{}])
        h0 = heteronyms[0] if heteronyms else {}

        return {
            "title": entry.get("title"),
            "tl": entry.get("tl"),
            "poj": entry.get("poj"),
            "bopomofo": h0.get("bopomofo"),
        }
    except Exception:
        return None


# ============================
#  萌典台語 API（以台羅輸入式查）
# ============================

def query_moedict_tl(tl_numeric: str):
    """用台羅輸入式（含數字調）查萌典台語詞目。查不到回 None。"""
    try:
        url = f"https://www.moedict.tw/t/{tl_numeric}.json"
        resp = requests.get(url, timeout=5)
        if resp.status_code != 200:
            return None
        return resp.json()
    except Exception:
        return None


# ============================
#  RIME 字典載入
# ============================

def load_rime_dict(path="tl_ji_khoo.dict.yaml"):
    """從 RIME 字典檔載入：漢字 → [台羅輸入式列表]。"""
    if not os.path.exists(path):
        print(f"找不到 RIME 字典檔：{path}")
        return {}

    local_dict = {}
    in_data = False

    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()

            # YAML 區塊結束，後面才是字典內容
            if line == "...":
                in_data = True
                continue

            if not in_data:
                continue

            if not line or line.startswith("#"):
                continue

            parts = line.split("\t")
            if len(parts) < 2:
                continue

            han = parts[0]
            tl = parts[1]  # 台羅輸入式（數字調）

            if han not in local_dict:
                local_dict[han] = []

            if tl not in local_dict[han]:
                local_dict[han].append(tl)

    return local_dict


# ============================
#  台羅處理：數字調 → 調號式
# ============================

TONE_MARK = {
    "1": "ˉ",   # 陰平（通常不寫，但先保留）
    "2": "ˊ",   # 陰上
    "3": "",    # 陰去（多平調，無符號）
    "4": "ˋ",   # 陰入
    "5": "ˇ",   # 陽上（簡化處理）
    "7": "",    # 陽去（多平調，無符號）
    "8": "̍",   # 陽入（用 a̍ 型）
}

# 這裡先用簡化版母音集合，實務上可再細修
VOWELS = "aeiouo͘"


def numeric_tl_to_marked(tl: str) -> str:
    """
    將台羅輸入式（數字調）轉為大致可用的調號式。
    例： tsioh8 → tsio̍h, peh8 → pe̍h
    """
    m = re.match(r"(.+?)([1-8])$", tl)
    if not m:
        return tl

    base, tone = m.group(1), m.group(2)
    mark = TONE_MARK.get(tone, "")

    # 若沒有調號（3、7），直接回 base
    if not mark:
        return base

    # 找第一個母音加上調號（簡化版規則）
    for i, ch in enumerate(base):
        if ch in VOWELS:
            return base[:i] + ch + mark + base[i + 1 :]

    return tl


# ============================
#  台羅 → POJ（非常簡化版）
# ============================

TL_TO_POJ_INITIAL = {
    "tsh": "chh",
    "ts": "ch",
    "j": "j",
    "ph": "ph",
    "th": "th",
    "kh": "kh",
    "p": "p",
    "t": "t",
    "k": "k",
    "m": "m",
    "n": "n",
    "l": "l",
    "s": "s",
    "h": "h",
    "g": "g",
    "b": "b",
}


def tl_to_poj(tl: str) -> str:
    """
    粗略將台羅輸入式轉為 POJ（白話字），目前只處理聲母與簡單韻母。
    用於 CLI 顯示參考，不追求 100% 嚴格。
    """
    m = re.match(r"(.+?)([1-8])$", tl)
    if not m:
        return tl

    base, tone = m.group(1), m.group(2)

    # 聲母
    initial = ""
    for ini in sorted(TL_TO_POJ_INITIAL, key=len, reverse=True):
        if base.startswith(ini):
            initial = TL_TO_POJ_INITIAL[ini]
            base = base[len(ini) :]
            break

    # 韻母簡單處理：oo → o͘
    poj_rime = base.replace("oo", "o͘")

    # 先合併，再用台羅的調號函式處理（偷懶但實用）
    combined = initial + poj_rime
    marked = numeric_tl_to_marked(combined[:-1] + tone) if combined else tl

    return marked


# ============================
#  台羅 → IPA（簡化版）
# ============================

TL_INITIAL_TO_IPA = {
    "tsh": "tsh",
    "ts": "ts",
    "j": "dz",
    "ph": "pʰ",
    "th": "tʰ",
    "kh": "kʰ",
    "p": "p",
    "t": "t",
    "k": "k",
    "m": "m",
    "n": "n",
    "l": "l",
    "s": "s",
    "h": "h",
    "g": "ɡ",
    "b": "b",
}

TL_FINAL_TO_IPA = {
    "a": "a",
    "ah": "aʔ",
    "ann": "ã",
    "i": "i",
    "iah": "iaʔ",
    "io": "io",
    "iu": "iu",
    "u": "u",
    "o": "o",
    "oh": "oʔ",
    "ong": "oŋ",
    "eng": "eŋ",
}


def tl_to_ipa(tl: str) -> str:
    """
    台羅輸入式 → 簡化版 IPA。主要为參考用，非完整音韻學版本。
    """
    m = re.match(r"(.+?)([1-8])$", tl)
    if not m:
        return tl

    base, tone = m.group(1), m.group(2)

    # 聲母
    initial_ipa = ""
    for ini in sorted(TL_INITIAL_TO_IPA, key=len, reverse=True):
        if base.startswith(ini):
            initial_ipa = TL_INITIAL_TO_IPA[ini]
            base = base[len(ini) :]
            break

    # 韻母
    final_ipa = TL_FINAL_TO_IPA.get(base, base)

    # 調號目前不反映在 IPA（可以之後加 tone sandhi 等）
    return f"{initial_ipa}{final_ipa}"


# ============================
#  三合一查詢（以 RIME 為主體）
# ============================

def smart_query(han: str, local_dict: dict):
    """
    三合一查詢邏輯：
    1. 教育部（若成功）
    2. RIME 字典（主體，必回）
    3. 萌典以台羅輸入式查補強
    """
    # ① 教育部（若未來解封）
    edu = query_edu(han)
    if edu:
        return {
            "source": "教育部",
            "han": edu.get("title", han),
            "readings": [
                {
                    "tl_numeric": edu.get("tl"),
                    "tl_marked": numeric_tl_to_marked(edu.get("tl")) if edu.get("tl") else None,
                    "poj": edu.get("poj"),
                    "ipa": None,
                    "moedict": None,
                }
            ],
        }

    # ② RIME 本地字典（主體）
    if han in local_dict:
        tls = local_dict[han]
        readings = []
        for tl in tls:
            tl_numeric = tl
            tl_marked = numeric_tl_to_marked(tl_numeric)
            poj = tl_to_poj(tl_numeric)
            ipa = tl_to_ipa(tl_numeric)
            moe = query_moedict_tl(tl_numeric)

            readings.append(
                {
                    "tl_numeric": tl_numeric,
                    "tl_marked": tl_marked,
                    "poj": poj,
                    "ipa": ipa,
                    "moedict": moe,
                }
            )

        return {
            "source": "RIME",
            "han": han,
            "readings": readings,
        }

    # ③ 若未來想支援「直接輸入台羅/POJ」，可在這裡加另一條路徑
    return None


# ============================
#  CLI 輸出
# ============================

def print_result(result: dict):
    print(f"來源：{result['source']}")
    print(f"漢字：{result['han']}")

    for idx, r in enumerate(result["readings"], start=1):
        print("\n--- 讀音", idx, "---")
        print(f"台羅（輸入式）：{r['tl_numeric']}")
        print(f"台羅（調號式）：{r['tl_marked']}")
        print(f"白話字（POJ）：{r['poj']}")
        print(f"IPA：{r['ipa']}")

        # 若有萌典補強資料
        moe = r.get("moedict")
        if moe:
            print("（萌典補強）")
            title = moe.get("title")
            het = moe.get("heteronyms", [{}])[0]
            print(f"  萌典詞目：{title}")
            print(f"  萌典 POJ：{het.get('poj')}")
            print(f"  萌典方音符號：{het.get('bopomofo')}")

        # ★★★ 在每一筆讀音後標示來源 ★★★
        if result["source"] == "RIME":
            print("來源：個人字典")
        elif result["source"] == "教育部":
            print("來源：教育部辭典")
        else:
            print("來源：其他")

# def print_result(result: dict):
#     print(f"來源：{result['source']}")
#     print(f"漢字：{result['han']}")

#     for idx, r in enumerate(result["readings"], start=1):
#         print("\n--- 讀音", idx, "---")
#         print(f"台羅（輸入式）：{r['tl_numeric']}")
#         print(f"台羅（調號式）：{r['tl_marked']}")
#         print(f"白話字（POJ）：{r['poj']}")
#         print(f"IPA：{r['ipa']}")

#         moe = r.get("moedict")
#         if moe:
#             print("（萌典補強）")
#             title = moe.get("title")
#             het = moe.get("heteronyms", [{}])[0]
#             print(f"  萌典詞目：{title}")
#             print(f"  萌典 POJ：{het.get('poj')}")
#             print(f"  萌典方音符號：{het.get('bopomofo')}")


# ============================
#  主程式
# ============================

def main():
    parser = argparse.ArgumentParser(
        description="台語字典 CLI（RIME + 台羅 + POJ + IPA + 萌典補強）"
    )
    parser.add_argument("word", help="要查的漢字，例如：白、石")
    parser.add_argument(
        "--dict",
        "-d",
        help="RIME 字典檔路徑（預設為 tl_ji_khoo.dict.yaml）",
        default="tl_ji_khoo.dict.yaml",
    )
    args = parser.parse_args()

    local_dict = load_rime_dict(args.dict)

    result = smart_query(args.word, local_dict)

    if not result:
        print("查無資料（RIME 字典中沒有這個漢字的讀音）")
        return

    print_result(result)


if __name__ == "__main__":
    main()
