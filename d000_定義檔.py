"""d000_定義檔.py v0.0.1

定義程式所需之常數、設定參數。

Dict_List 為批次作業清單：d100／d200 會依序處理其中每一筆。
各筆可選欄位：
  - 標的工作表名稱：d100 轉換後工作表名（預設「台語注音二式」）
  - 標音欄：漢字標音所在欄（預設「B」；可用 d100 --col_addr 全域覆寫）
"""

from __future__ import annotations

import os
from pathlib import Path

# ---------------------------------------------------------------------
# 定義 RIME 字典檔的標頭內容（Header）樣板
# 以 str.format(...) 填入：字典檔來源摘要、漢字標音系統名稱、字典簡介、
# 輸入方案名稱、版本號。
# ---------------------------------------------------------------------
rime_header = """# Rime dictionary
# encoding: utf-8
#
# {字典檔來源摘要}：{漢字標音系統名稱}
# {字典簡介}
#
---
name: {輸入方案名稱}
version: "{版本號}"
sort: by_weight
use_preset_vocabulary: false
columns:
  - text    # 漢字
  - code    # 漢字讀音標音
  - weight  # 常用度（優先顯示度）
  - stem    # 用法舉例
  - create  # 建立時間
# import_tables:
#   - ji_khoo_ziann_ji_bpm2
#   - ji_khoo_ban_lam_bpm2
#   - ji_khoo_su_lui_bpm2
...
"""

source_dir_path = r"%UserProfile%\work\rime-tlpa\src"
target_dir_path = r"%UserProfile%\work\rime-tlpa\src"

# 預設：來源「台羅拼音」→ 標的「台語注音二式」；標音在 B 欄（code）
Dict_List = {
    "HanJi": {
        # 【漢字庫】內「台羅拼音」為資料庫欄位格式；RIME 用「正字庫*」工作表
        "WorkBook檔案": "【漢字庫】.xlsx",
        "WorkSheet名稱": "正字庫台羅拼音",
        "標的工作表名稱": "正字庫台語注音二式",
        "字典檔來源摘要": "網路蒐集之《漢字庫》",
        "字典簡介": "閩南語之漢字字典檔（含：文讀音、白話音）",
        "版本號": "v0.1.0",
        "TL": {
            "輸入方案名稱": "ji_khoo_tl_HanJi",
        },
        "BPM2": {
            "輸入方案名稱": "ji_khoo_bpm2_HanJi",
        },
    },
    "BanLam": {
        "WorkBook檔案": "【BanLam字典】.xlsx",
        "WorkSheet名稱": "台羅拼音",
        "字典檔來源摘要": "網路蒐集之《BanLam 字典》",
        "字典簡介": "閩南語漢字之單字字典，含：各地閩南話方音",
        "版本號": "v0.1.0",
        "TL": {
            "輸入方案名稱": "ji_khoo_tl_BanLam",
        },
        "BPM2": {
            "輸入方案名稱": "ji_khoo_bpm2_BanLam",
        },
    },
    "KamJiTian": {
        "WorkBook檔案": "【甘字典】ChhoeTaigi_KamJitian.xlsx",
        "WorkSheet名稱": "台羅拼音",
        "字典檔來源摘要": "《甘字典》",
        "字典簡介": "漢字輸入用字典檔",
        "版本號": "v0.1.0",
        "TL": {
            "輸入方案名稱": "ji_khoo_tl_KamJiTian",
        },
        "BPM2": {
            "輸入方案名稱": "ji_khoo_bpm2_KamJiTian",
        },
    },
    "SuLui": {
        "WorkBook檔案": "【1956台灣白話基礎語句】ChhoeTaigi_TaioanPehoeKichhooGiku.xlsx",
        "WorkSheet名稱": "台羅拼音",
        "字典檔來源摘要": "《1956台灣白話基礎語句》",
        "字典簡介": "辭彙輸入用字典檔",
        "版本號": "v0.1.0",
        "TL": {
            "輸入方案名稱": "ji_khoo_tl_su_lui",
        },
        "BPM2": {
            "輸入方案名稱": "ji_khoo_bpm2_su_lui",
        },
    },
    "ZiannJi": {
        # 實際檔名含【後空白，且工作表為 RIME_Dict
        "WorkBook檔案": "【 閩南語白話音漢字正字】.xlsx",
        "WorkSheet名稱": "RIME_Dict",
        "字典檔來源摘要": "網路蒐集之《閩南語白話漢字正字》",
        "字典簡介": "閩南語漢字之單字/辭彙輸入方案用字典檔",
        "版本號": "v0.1.0",
        "TL": {
            "輸入方案名稱": "ji_khoo_tl_ziann_ji",
        },
        "BPM2": {
            "輸入方案名稱": "ji_khoo_bpm2_ziann_ji",
        },
    },
}

DEFAULT_SOURCE_SHEET = "台羅拼音"
DEFAULT_TARGET_SHEET = "台語注音二式"
DEFAULT_CODE_COL = "B"


def expand_dir_path(path: str) -> Path:
    """展開環境變數（如 %UserProfile%）並回傳 Path。"""
    return Path(os.path.expandvars(path)).expanduser()


def get_source_sheet_name(dict_cfg: dict) -> str:
    return dict_cfg.get("WorkSheet名稱") or DEFAULT_SOURCE_SHEET


def get_target_sheet_name(dict_cfg: dict) -> str:
    return dict_cfg.get("標的工作表名稱") or DEFAULT_TARGET_SHEET


def get_code_col_addr(dict_cfg: dict) -> str:
    return dict_cfg.get("標音欄") or DEFAULT_CODE_COL


def format_rime_header(
    dict_cfg: dict,
    *,
    輸入方案名稱: str,
    漢字標音系統名稱: str,
) -> str:
    """依字典設定與音標系統名稱，組出 RIME 字典檔標頭。"""
    return rime_header.format(
        字典檔來源摘要=dict_cfg.get("字典檔來源摘要", ""),
        漢字標音系統名稱=漢字標音系統名稱,
        字典簡介=dict_cfg.get("字典簡介", ""),
        輸入方案名稱=輸入方案名稱,
        版本號=dict_cfg.get("版本號", "v0.1.0"),
    )
