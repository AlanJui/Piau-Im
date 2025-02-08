import os
import sys
from datetime import datetime

import pandas as pd
import win32com.client  # 用於獲取作用中的 Excel 檔案
import yaml


def get_active_excel_file():
    """
    獲取當前作用中的 Excel 檔案路徑。
    如果沒有作用中的 Excel 檔案，返回 None。
    """
    try:
        # 獲取 Excel 應用程式
        excel_app = win32com.client.GetObject(Class="Excel.Application")
        if excel_app is None:
            print("❌ 沒有作用中的 Excel 檔案。")
            return None

        # 獲取作用中的工作簿
        active_workbook = excel_app.ActiveWorkbook
        if active_workbook is None:
            print("❌ 沒有作用中的 Excel 工作簿。")
            return None

        # 獲取檔案路徑
        excel_file = active_workbook.FullName
        print(f"✅ 作用中的 Excel 檔案：{excel_file}")
        return excel_file

    except Exception as e:
        print(f"❌ 獲取作用中的 Excel 檔案失敗: {e}")
        return None

def export_to_rime_dictionary():
    try:
        # 獲取作用中的 Excel 檔案路徑
        EXCEL_FILE = get_active_excel_file()
        if EXCEL_FILE is None:
            return 1  # 退出程式

        # 讀取 Excel 工作表
        SHEET_NAME = '甲骨釋文漢字庫'  # 替換為你的工作表名稱
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

        # 設定字典檔內容
        rime_dict = {
            "name": "tl_ji_khoo_peh_ue",
            "version": "v0.1.0.0",
            "sort": "by_weight",
            "use_preset_vocabulary": False,
            "columns": [
                {"text": "漢字"},
                {"code": "台灣音標（TLPA）拼音"},
                {"weight": "常用度（優先顯示度）"},
                {"stem": "用法舉例"},
                {"create": "建立時間"}
            ],
            "import_tables": [
                "tl_ji_khoo_kah_kut_bun"
            ],
            "entries": []
        }

        # 將 Excel 資料轉換為 RIME 字典格式
        for index, row in df.iterrows():
            entry = {
                "text": row["漢字"],
                "code": row["台羅音標"],
                "weight": row["常用度"],
                "stem": row["摘要說明"] if pd.notna(row["摘要說明"]) else "",
                "create": row["更新時間"].strftime("%Y/%m/%d %H:%M") if pd.notna(row["更新時間"]) else ""
            }
            rime_dict["entries"].append(entry)

        # 設定輸出字典檔名稱(YAML檔)
        DICT_FILENAME = 'tl_ji_khoo_kah_kut_bun.dict.yaml'
        with open(DICT_FILENAME, 'w', encoding='utf-8') as file:
            yaml.dump(rime_dict, file, allow_unicode=True, sort_keys=False)

        print(f"✅ RIME 字典 `{DICT_FILENAME}` 匯出完成！")
        return 0

    except Exception as e:
        print(f"❌ 匯出 RIME 字典失敗: {e}")
        return 1

if __name__ == "__main__":
    exit_code = export_to_rime_dictionary()
    sys.exit(exit_code)