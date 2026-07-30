"""d000_定義檔.py v0.0.1 v0.0.1

定義程式所需之常數、設定參數。
"""

# ---------------------------------------------------------------------
# 定義 RIME 字典檔的標頭內容（Header）
# ---------------------------------------------------------------------
rime_header = f"""# Rime dictionary
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

source_dir_path = "%UserProfile%\work\rime-tlpa\src"
target_dir_path = "%UserProfile%\work\rime-tlpa"

dict_list = {
    "ji_khoo_tl_HanJi": {
        "WorkBook檔案": "【漢字庫】.xlsx",
        "WorkSheet名稱": "台羅拼音",
        "字典檔來源摘要": "網路蒐集",
        "漢字標音系統名稱": "台羅拼音",
        "字典簡介": "閩南語之漢字字典檔（含：文讀音、白話音）",
        "輸入方案名稱": "ji_khoo_tl_HanJi",
        "版本號": "v0.1.0",
    },
    "ji_khoo_bpm2_HanJi": {
        "WorkBook檔案": "【漢字庫】.xlsx",
        "WorkSheet名稱": "台語注音二式",
        "字典檔來源摘要": "網路蒐集",
        "漢字標音系統名稱": "台語注音二式",
        "字典簡介": "閩南語之漢字字典檔（含：文讀音、白話音）",
        "輸入方案名稱": "ji_khoo_bpm2_HanJi",
        "版本號": "v0.1.0",
    },
    "ji_khoo_tl_BanLam": {
        "WorkBook檔案": "【BanLam字典】.xlsx",
        "WorkSheet名稱": "台羅拼音",
        "字典檔來源摘要": "網路蒐集",
        "漢字標音系統名稱": "台羅拼音",
        "字典簡介": "閩南語漢字之單字字典，含：各地閩南話方音",
        "輸入方案名稱": "ji_khoo_tl_BanLam",
        "版本號": "v0.1.0",
    },
    "ji_khoo_bpm2_BanLam": {
        "WorkBook檔案": "【BanLam字典】.xlsx",
        "WorkSheet名稱": "台語注音二式",
        "字典檔來源摘要": "網路蒐集",
        "漢字標音系統名稱": "台語注音二式",
        "字典簡介": "閩南語漢字之單字字典，含：各地閩南話方音",
        "輸入方案名稱": "ji_khoo_bpm2_BanLam",
        "版本號": "v0.1.0",
    },
    "ji_khoo_tl_KamJiTian": {
        "WorkBook檔案": "【甘字典】ChhoeTaigi_KamJitian.xlsx",
        "WorkSheet名稱": "台羅拼音",
        "字典檔來源摘要": "《甘字典》",
        "漢字標音系統名稱": "台羅拼音",
        "字典簡介": "漢字輸入用字典檔",
        "輸入方案名稱": "ji_khoo_tl_KamJiTian",
        "版本號": "v0.1.0",
    },
    "ji_khoo_bpm2_KamJiTian": {
        "WorkBook檔案": "【甘字典】ChhoeTaigi_KamJitian.xlsx",
        "WorkSheet名稱": "台語注音二式",
        "字典檔來源摘要": "《甘字典》",
        "漢字標音系統名稱": "台語注音二式",
        "字典簡介": "漢字輸入用字典檔",
        "輸入方案名稱": "ji_khoo_bpm2_KamJiTian",
        "版本號": "v0.1.0",
    },
    "ji_khoo_tl_ziann_ji": {
        "字典檔來源摘要": "網路蒐集",
        "漢字標音系統名稱": "台羅拼音",
        "字典簡介": "閩南語漢字之單字/辭彙輸入方案用字典檔",
        "輸入方案名稱": "ji_khoo_tl_ziann_ji",
        "版本號": "v0.1.0",
    },
    "ji_khoo_bpm2_ziann_ji": {
        "字典檔來源摘要": "網路蒐集",
        "漢字標音系統名稱": "台語注音二式",
        "字典簡介": "閩南語漢字之單字/辭彙輸入方案用字典檔",
        "輸入方案名稱": "ji_khoo_bpm2_ziann_ji",
        "版本號": "v0.1.0",
    },
    "ji_khoo_tl_su_lui": {
        "WorkBook檔案": "【1956台灣白話基礎語句】ChhoeTaigi_TaioanPehoeKichhooGiku.xlsx",
        "漢字標音系統名稱": "台羅拼音",
        "字典簡介": "辭彙輸入用字典檔",
        "輸入方案名稱": "ji_khoo_tl_su_lui",
        "版本號": "v0.1.0",
    },
    "ji_khoo_bpm2_su_lui": {
        "WorkBook檔案": "【1956台灣白話基礎語句】ChhoeTaigi_TaioanPehoeKichhooGiku.xlsx",
        "WorkSheet名稱": "台語注音二式",
        "字典檔來源摘要": "《1956台灣白話基礎語句》",
        "漢字標音系統名稱": "台語注音二式",
        "字典簡介": "辭彙輸入用字典檔",
        "輸入方案名稱": "ji_khoo_bpm2_su_lui",
        "版本號": "v0.1.0",
    },
}