"""
mod_han_ji_tian_api.py

呼叫三合一漢字典（han-ji-tian）讀音 API，轉成與 HanJiTian.han_ji_ca_piau_im
相同形狀的選項清單，供 a251／a300 選音後寫回工作表。

環境變數：
  HAN_JI_TIAN_API_URL   預設 https://han-ji-tian.apps.grok.bot
  HAN_JI_TIAN_API_TOKEN 或 READING_API_TOKEN（已設定時帶 Bearer）
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional, Union

from mod_標音 import split_tai_gi_im_piau

DEFAULT_API_BASE = "https://han-ji-tian.apps.grok.bot"


def _api_base() -> str:
    return (os.environ.get("HAN_JI_TIAN_API_URL") or DEFAULT_API_BASE).rstrip("/")


def _api_token() -> str:
    return (
        os.environ.get("HAN_JI_TIAN_API_TOKEN")
        or os.environ.get("READING_API_TOKEN")
        or ""
    ).strip()


def fetch_reading_json(han_ji: str, timeout: float = 20.0) -> Dict[str, Any]:
    """GET /api/v1/reading?han=…，回傳原始 JSON。"""
    han = (han_ji or "").strip()
    if not han:
        raise ValueError("漢字不可為空白")

    url = f"{_api_base()}/api/v1/reading?han={urllib.parse.quote(han)}"
    headers = {"Accept": "application/json", "User-Agent": "Piau-Im-a251/1.0"}
    token = _api_token()
    if token:
        headers["Authorization"] = f"Bearer {token}"

    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body)
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"三合一字典 API HTTP {e.code}：{detail}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"無法連線三合一字典 API：{e}") from e


def _put_reading(
    bag: Dict[str, Dict[str, Union[str, float]]],
    han_ji: str,
    tai_lo: str,
    freq: float,
    note: str,
) -> None:
    key = (tai_lo or "").strip().lower()
    if not key:
        return
    prev = bag.get(key)
    if prev is None or float(prev.get("常用度") or 0) < freq:
        bag[key] = {
            "識別號": 0,
            "漢字": han_ji,
            "台語音標": key,
            "常用度": freq,
            "摘要說明": note,
        }
    elif note and note not in str(prev.get("摘要說明") or ""):
        prev["摘要說明"] = f"{prev.get('摘要說明') or ''}｜{note}".strip("｜")


def readings_to_entries(
    payload: Dict[str, Any],
) -> Optional[List[Dict[str, Union[str, float]]]]:
    """
    將 API-001 JSON 壓成 HanJiTian 選項清單（含聲／韻／調拆解）。
    常用度啟發式：雅俗通 0.9、甘文 0.8、甘白 0.6、廣韻 0.5。
    """
    han_ji = str(payload.get("han_ji") or "").strip()
    if not han_ji:
        return None

    bag: Dict[str, Dict[str, Union[str, float]]] = {}

    nga = payload.get("nga_siok_thong") or {}
    if isinstance(nga, dict):
        for tai_lo, vals in nga.items():
            fifteen = ""
            if isinstance(vals, list) and vals:
                fifteen = str(vals[0] or "")
            _put_reading(bag, han_ji, tai_lo, 0.9, f"雅俗通 {fifteen}".strip())

    kam = payload.get("kam_ji_tian") or {}
    if isinstance(kam, dict):
        for tai_lo, vals in kam.items():
            mode = ""
            meaning = ""
            if isinstance(vals, list):
                if len(vals) >= 2:
                    mode = str(vals[1] or "")
                if len(vals) >= 3:
                    meaning = str(vals[2] or "")
                elif len(vals) >= 1:
                    meaning = str(vals[0] or "")
            freq = 0.8 if mode == "文" else 0.6
            note = f"甘字典{mode} {meaning}".strip()
            _put_reading(bag, han_ji, tai_lo, freq, note)

    kong = payload.get("kong_un") or {}
    if isinstance(kong, dict):
        for tai_lo, vals in kong.items():
            cut = ""
            meaning = ""
            if isinstance(vals, list):
                if len(vals) >= 1:
                    cut = str(vals[0] or "")
                if len(vals) >= 2:
                    meaning = str(vals[1] or "")[:40]
            _put_reading(bag, han_ji, tai_lo, 0.5, f"廣韻 {cut} {meaning}".strip())

    if not bag:
        return None

    data: List[Dict[str, Union[str, float]]] = []
    for entry in sorted(
        bag.values(),
        key=lambda e: float(e.get("常用度") or 0),
        reverse=True,
    ):
        tai_loo_im = str(entry.get("台語音標") or "")
        try:
            split_result = split_tai_gi_im_piau(tai_loo_im)
        except (TypeError, IndexError, ValueError):
            continue
        if not split_result or not any(split_result):
            continue
        entry["聲母"] = split_result[0]
        entry["韻母"] = split_result[1]
        entry["聲調"] = split_result[2]
        entry["台語音標"] = f"{entry['聲母']}{entry['韻母']}{entry['聲調']}"
        data.append(entry)

    return data or None


def han_ji_ca_piau_im_from_api(
    han_ji: str,
    ue_im_lui_piat: str = "文讀音",
    display_all_piau_im: bool = True,
    tai_lo_im_piau: Optional[str] = None,
) -> Optional[List[Dict[str, Union[str, float]]]]:
    """對齊 HanJiTian.han_ji_ca_piau_im 簽名，供 a251 置換查詢來源。"""
    del ue_im_lui_piat, display_all_piau_im  # API 一次回三庫，不依文白篩
    payload = fetch_reading_json(han_ji)
    entries = readings_to_entries(payload)
    if not entries:
        return None
    if tai_lo_im_piau:
        want = tai_lo_im_piau.strip().lower()
        filtered = [
            e
            for e in entries
            if str(e.get("台語音標") or "").lower() == want
            or str(e.get("台語音標") or "").replace("z", "ts").lower() == want
        ]
        return filtered or None
    return entries
