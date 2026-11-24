"""
mod_han_ji_tian_api.py

呼叫三合一漢字典（han-ji-tian）讀音 API，轉成與 HanJiTian.han_ji_ca_piau_im
相同形狀的選項清單，供 a251／a300 選音後寫回工作表。

憑證檔（專案根目錄 `.api_key`，已列入 .gitignore）：
  HAN_JI_TIAN_API_TOKEN=…
  HAN_JI_TIAN_API_URL=https://han-ji-tian.apps.grok.bot   # 可選

亦可改用系統環境變數 HAN_JI_TIAN_API_TOKEN／READING_API_TOKEN。
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from mod_標音 import split_tai_gi_im_piau

DEFAULT_API_BASE = "https://han-ji-tian.apps.grok.bot"
_API_KEY_LOADED = False


def load_api_key_file(override: bool = False) -> Path | None:
    """載入專案根目錄 `.api_key` 至環境變數；檔案不存在則略過。"""
    global _API_KEY_LOADED
    path = Path(__file__).resolve().parent / ".api_key"
    if not path.is_file():
        return None
    try:
        from dotenv import load_dotenv

        load_dotenv(path, override=override)
    except ImportError:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key, val = key.strip(), val.strip().strip("'").strip('"')
            if key and (override or key not in os.environ):
                os.environ[key] = val
    _API_KEY_LOADED = True
    return path


def _ensure_api_key_loaded() -> None:
    if not _API_KEY_LOADED:
        load_api_key_file(override=False)


def _api_base() -> str:
    _ensure_api_key_loaded()
    return (os.environ.get("HAN_JI_TIAN_API_URL") or DEFAULT_API_BASE).rstrip("/")


def _api_token() -> str:
    _ensure_api_key_loaded()
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


# 顯示與去重優先序：雅俗通 → 甘字典 → 廣韻
SOURCE_ORDER = ("雅俗通", "甘字典", "廣韻")
SOURCE_RANK = {name: i for i, name in enumerate(SOURCE_ORDER)}


def _put_reading(
    bag: Dict[str, Dict[str, Union[str, float]]],
    han_ji: str,
    tai_lo: str,
    freq: float,
    note: str,
    source: str,
    wen_bai: str = "",
) -> None:
    """同台羅只保留先寫入者（來源優先序由呼叫端保證）。"""
    key = (tai_lo or "").strip().lower()
    if not key or key in bag:
        return
    bag[key] = {
        "識別號": 0,
        "漢字": han_ji,
        "台語音標": key,
        "常用度": freq,
        "摘要說明": note,
        "資料來源": source,
        "文白": wen_bai if wen_bai in ("文", "白") else "",
    }


def _finalize_entry(
    entry: Dict[str, Union[str, float]],
) -> Optional[Dict[str, Union[str, float]]]:
    tai_loo_im = str(entry.get("台語音標") or "")
    try:
        split_result = split_tai_gi_im_piau(tai_loo_im)
    except (TypeError, IndexError, ValueError):
        return None
    if not split_result or not any(split_result):
        return None
    entry["聲母"] = split_result[0]
    entry["韻母"] = split_result[1]
    entry["聲調"] = split_result[2]
    entry["台語音標"] = f"{entry['聲母']}{entry['韻母']}{entry['聲調']}"
    return entry


def readings_to_entries(
    payload: Dict[str, Any],
) -> Optional[List[Dict[str, Union[str, float]]]]:
    """
    將 API-001 JSON 壓成 HanJiTian 選項清單（含聲／韻／調拆解）。
    同音多來源時歸屬優先：雅俗通 → 甘字典 → 廣韻。
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
            _put_reading(
                bag,
                han_ji,
                tai_lo,
                0.9,
                f"雅俗通 {fifteen}".strip(),
                "雅俗通",
            )

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
            _put_reading(
                bag, han_ji, tai_lo, freq, note, "甘字典", wen_bai=mode
            )

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
            _put_reading(
                bag,
                han_ji,
                tai_lo,
                0.5,
                f"廣韻 {cut} {meaning}".strip(),
                "廣韻",
            )

    if not bag:
        return None

    # 穩定排序：僅依資料來源分組，組內維持 API／寫入順序
    data: List[Dict[str, Union[str, float]]] = []
    for entry in sorted(
        bag.values(),
        key=lambda e: SOURCE_RANK.get(str(e.get("資料來源") or ""), 99),
    ):
        finalized = _finalize_entry(entry)
        if finalized:
            data.append(finalized)

    return data or None


def format_grouped_reading_lines(
    han_ji: str,
    result: List[Dict[str, Union[str, float]]],
    piau_im_options: List[tuple[str, str]],
) -> None:
    """在 Console 依《雅俗通》《甘字典》《廣韻》分組列出選項（編號連號）。"""
    print(f"【{han_ji}】有 {len(result)} 個讀音：")
    last_source = None
    for idx, entry in enumerate(result):
        source = str(entry.get("資料來源") or "其它")
        if source != last_source:
            print(f"《{source}》：")
            last_source = source
        tai_gi_im_piau, han_ji_piau_im = piau_im_options[idx]
        wen_bai = str(entry.get("文白") or "").strip()
        suffix = f"  【{wen_bai}】" if source == "甘字典" and wen_bai in ("文", "白") else ""
        print(
            f"  {idx + 1}. 台語音標：{tai_gi_im_piau:<8} "
            f"漢字標音：[{han_ji_piau_im}]{suffix}"
        )


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
