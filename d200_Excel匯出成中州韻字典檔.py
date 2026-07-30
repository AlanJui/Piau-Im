""" d200_Excel匯出成中州韻字典檔.py v0.0.1
摘要：
    將 Excel 檔案中【工作表】的資料轉換成中州韻字典檔。
    1. 打開 Excel 檔案，讀取【台羅拼音工作表】/【台語注音二式工作表】的【字典檔內容】（A欄：E欄）資料。
    2. 套用【Rime 字典檔的標頭內容】及【字典檔內容】，製作成【中州韻字典檔】。
    3. 將【中州韻字典檔】存放在【target_dir_path】目錄路徑下。

指令：
    python d200_Excel匯出成中州韻字典檔.py

參考：
    - a860_將Excel檔中的台羅拼音轉換成台語注音二式.py
"""  # noqa: N999

import os

from d000_定義檔 import (
    Dict_List,
    source_dir_path,
    target_dir_path,
)

ji_khoo = Dict_List["HanJi"]
# 定義 Excel 檔案名稱，及使用工作表名稱
excel_file_name = ji_khoo["WorkBook檔案"]
excel_sheet_name = ji_khoo["WorkSheet名稱"]
excel_file_path = os.path.join(
    source_dir_path,
    ji_khoo["WorkBook檔案"],
)

# 定義字典檔檔名
dict_name_tl = ji_khoo["TL"]["輸入方案名稱"]
dict_name_bpm2 = ji_khoo["BPM2"]["輸入方案名稱"]

dict_name = dict_name_tl
dict_file_name = f"{dict_name}.dict.yaml"
dict_file_path = os.path.join(
    target_dir_path,
    dict_file_name,
)

# ---------------------------------------------------------------------
# 定義 RIME 字典檔的標頭內容（Header）
# ---------------------------------------------------------------------
rime_header_content = f"""# Rime dictionary
# encoding: utf-8
#
# {ji_khoo["字典檔來源摘要"]}：{ji_khoo["漢字標音系統名稱"]}
# {ji_khoo["字典簡介"]}
#
---
name: {dict_name}
version: "{ji_khoo["版本號"]}"
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