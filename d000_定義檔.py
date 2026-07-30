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

source_dir_path = "%UserProfile%\\work\\rime-tlpa\\src"
target_dir_path = "%UserProfile%\\work\\rime-tlpa\\src"

Dict_List = {
    "HanJi": {
        "WorkBook檔案": "【漢字庫】.xlsx",
        "WorkSheet名稱": "台羅拼音",
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
        "TL": {
            "輸入方案名稱": "ji_khoo_tl_su_lui",
        },
        "BPM2": {
            "輸入方案名稱": "ji_khoo_bpm2_su_lui",
        },
    },
    "ZiannJi": {
        "WorkBook檔案": "【閩南語白話漢字正字】.xlsx",
        "WorkSheet名稱": "台羅拼音",
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