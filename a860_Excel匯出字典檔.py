# C:\Users\AlanJui\AppData\Roaming\Rime
import argparse
import sys
from datetime import datetime

import pandas as pd

from mod_excel_access import get_active_excel_file


def export_to_rime_dictionary(sheet_name, dict_filename, is_master_file):
    """
    將指定工作表轉換為 RIME 字典檔。
    :param sheet_name: Excel 工作表名稱
    :param dict_filename: 輸出的字典檔名稱
    :param is_master_file: 是否為母檔（True：母檔，False：子檔）
    """
    try:
        # 獲取作用中的 Excel 檔案路徑
        EXCEL_FILE = get_active_excel_file()
        if EXCEL_FILE is None:
            return 1  # 退出程式

        # 讀取 Excel 工作表
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)

        # 寫入 YAML 文件
        with open(dict_filename, 'w', encoding='utf-8') as file:
            # 寫入 RIME 字典檔頭
            file.write("# Rime dictionary\n")
            file.write("# encoding: utf-8\n")
            file.write("#\n")
            if is_master_file:
                file.write("# 河洛白話音\n")
            else:
                file.write("# 甲骨漢字庫\n")
                file.write("# 漢字讀者：白話音\n")
                file.write("# 漢字標音：使用【台語音標（TLPA）】\n")
            file.write("---\n")
            file.write(f"name: {dict_filename.replace('.dict.yaml', '')}\n")
            file.write('version: "0.1.0.0"\n')
            file.write("sort: by_weight\n")
            file.write("use_preset_vocabulary: false\n")
            file.write("columns:\n")
            file.write("  - text    # 漢字／詞彙\n")
            file.write("  - code    # 台灣音標（TLPA）拼音字母\n")
            file.write("  - weight  # 常用度（優先顯示度）\n")
            file.write("  - stem    # 用法舉例\n")
            file.write("  - create  # 建立日期\n")

            # 如果是母檔，加入 import_tables
            if is_master_file:
                file.write("import_tables:\n")
                file.write("  - tl_ji_khoo_kah_kut_bun      # 甲骨文考證漢字庫\n")
                # 其他可選的 import_tables（註解狀態）
                file.write("  # - tl_ji_khoo_peh_ue_cu_ting      # 個人自訂擴充字庫\n")
                file.write("  # - tl_ji_khoo_ciann_ji\n")
                file.write("  # - tl_ji_khoo_siong_iong_si_lui\n")
                file.write("  # - tl_ji_khoo_tai_uan_si_lui\n")

            file.write("...\n")

            # 寫入字典內容
            for index, row in df.iterrows():
                han_ji = row["漢字"]
                tai_lo_im_piau = row["台羅音標"]
                weight = row["常用度"]
                summary = row["摘要說明"] if pd.notna(row["摘要說明"]) else ""
                create_time = row["更新時間"].strftime("%Y-%m-%d %H:%M:%S") if pd.notna(row["更新時間"]) else ""

                # 將欄位之間以 Tab 分隔
                file.write(f"{han_ji}\t{tai_lo_im_piau}\t{weight}\t{summary}\t{create_time}\n")

        print(f"✅ RIME 字典 `{dict_filename}` 匯出完成！")
        return 0

    except Exception as e:
        print(f"❌ 匯出 RIME 字典失敗: {e}")
        return 1

def main():
    # 設定命令列參數
    parser = argparse.ArgumentParser(description="將 Excel 工作表轉換為 RIME 字典檔")
    parser.add_argument(
        "--kah-kut-bun",
        action="store_true",
        help="將【甲骨釋文漢字庫】工作表轉換為 RIME 字典檔（子檔）"
    )
    args = parser.parse_args()

    # 根據參數選擇工作表與字典檔名稱
    if args.kah_kut_bun:
        sheet_name = "甲骨釋文漢字庫"
        dict_filename = "tl_ji_khoo_kah_kut_bun.dict.yaml"
        is_master_file = False  # 子檔
    else:
        sheet_name = "漢字庫"
        dict_filename = "tl_ji_khoo_peh_ue.dict.yaml"
        is_master_file = True  # 母檔

    # 執行匯出
    exit_code = export_to_rime_dictionary(sheet_name, dict_filename, is_master_file)
    sys.exit(exit_code)

if __name__ == "__main__":
    main()