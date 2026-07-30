""" d100_Excel工作表自TL轉BPM2.py v0.0.1
摘要：
    將 Excel 檔案中的【台羅拼音】工作表轉換成【台語注音二式】工作表。
    1. 自【台羅拼音】工作表複製成【台語注音二式】工作表。
    2. 【台羅拼音】工作表，漢字標音欄預設【欄位址】為：B Column，可用 --col_addr 參數指定其他欄位址。
    3. 將【台羅拼音工作表】的漢字標音欄位，轉換成【台語注音二式】，置入【台語注音二式】工作表的漢字標音欄位。

指令：
    python d100_Excel工作表自TL轉BPM2.py --col_addr=[欄位址]

使用方法：
    1. 來源檔目錄路徑： %UserProfile%\\work\\rime-tlpa\\src
    2. 標的檔目錄路徑： %UserProfile%\\work\\rime-tlpa\\src
    3. Excel 檔案：Dict_List[index]["WorkBook檔案"]，如：【漢字庫】.xlsx
    4. 來源工作表名稱：Dict_List[index]["WorkSheet名稱"]，如：台羅拼音
    5. 標的工作表名稱：台語注音二式
"""  # noqa: N999

import os

from d000_定義檔 import (
    Dict_List,
    source_dir_path,
)

ji_khoo = Dict_List["HanJi"]
# 定義 Excel 檔案名稱，及使用工作表名稱
excel_file_name = os.path.join(
    source_dir_path,
    ji_khoo["WorkBook檔案"],
)

source_sheet_name = "台羅拼音"
target_sheet_name = "台語注音二式"