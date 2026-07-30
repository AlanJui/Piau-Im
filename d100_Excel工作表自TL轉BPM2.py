"""d100_Excel工作表自TL轉BPM2.py v0.0.1
摘要：
    將 Excel 檔案中的【台羅拼音】工作表轉換成【台語注音二式】工作表。

指令：
    python D100_Excel工作表自TL轉BPM2.py <欄位址>

使用方法：
    1. 來源檔目錄路徑： %UserProfile%\work\rime-tlpa\src
    2. 執行程式，程式會將 Excel 檔案中的資料轉換成中州韻字典檔。
    3. 中州韻字典檔會存放在程式所在的目錄下。
"""

import os

from d000_定義檔 import (
    dict_list,
    source_dir_path,
    target_dir_path,
)

ji_khoo_name = "ji_khoo_tl_HanJi"
# 定義 Excel 檔案名稱，及使用工作表名稱
excel_file_name = dict_list[ji_khoo_name]["WorkBook檔案"]
excel_sheet_name = dict_list[ji_khoo_name]["WorkSheet名稱"]
excel_file_path = os.path.join(
    source_dir_path,
    dict_list[ji_khoo_name]["WorkBook檔案"],
)

# 定義字典檔檔名
dict_name = dict_list[ji_khoo_name]["輸入方案名稱"]
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
# {dict_list[ji_khoo_name]["字典檔來源摘要"]}：{dict_list[ji_khoo_name]["漢字標音系統名稱"]}
# {dict_list[ji_khoo_name]["字典簡介"]}
#
---
name: {dict_list[ji_khoo_name]["輸入方案名稱"]}
version: "{dict_list[ji_khoo_name]["版本號"]}"
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