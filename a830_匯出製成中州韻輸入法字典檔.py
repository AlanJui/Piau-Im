# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import logging
import os
import sqlite3
import sys
from datetime import datetime

from dotenv import load_dotenv

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')

# =========================================================================
# 設定日誌
# =========================================================================
logging.basicConfig(
    filename='process_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# =========================================================================
# 常數定義
# =========================================================================
EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_INVALID_INPUT = 2
EXIT_CODE_PROCESS_FAILURE = 3
EXIT_CODE_UNKNOWN_ERROR = 99

# =========================================================================
# 功能 3：匯出成 RIME 輸入法之字典檔
# =========================================================================
def export_to_rime_dictionary():
    """
    將【漢字庫】資料表的資料匯出成 RIME 輸入法專用的字典檔（YAML 格式）。
    """
    conn = sqlite3.connect(DB_HO_LOK_UE)
    cursor = conn.cursor()

    try:
        # 讀取資料庫內容
        cursor.execute("SELECT 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫;")
        rows = cursor.fetchall()

        # 設定字典檔名稱
        dict_filename = "tl_ji_khoo_peh_ue.dict.yaml"

        # 寫入字典檔
        with open(dict_filename, "w", encoding="utf-8") as file:
            # 寫入 RIME 字典檔頭
            file.write("# Rime dictionary\n")
            file.write("# encoding: utf-8\n")
            file.write("#\n# 河洛白話音\n#\n")
            file.write("---\n")
            file.write("name: tl_ji_khoo_peh_ue\n")
            file.write("version: \"v0.1.0.0\"\n")
            file.write("sort: by_weight\n")
            file.write("use_preset_vocabulary: false\n")
            file.write("columns:\n")
            file.write("  - text    # 漢字\n")
            file.write("  - code    # 台灣音標（TLPA）拼音\n")
            file.write("  - weight  # 常用度（優先顯示度）\n")
            file.write("  - stem    # 用法舉例\n")
            file.write("  - create  # 建立時間\n")
            file.write("import_tables:\n")
            file.write("  - tl_ji_khoo_kah_kut_bun\n")
            file.write("...\n")

            # 寫入字典內容
            for han_ji, tai_lo_im_piau, weight, summary, create_time in rows:
                # 將欄位之間以 <tab> 分隔
                file.write(f"{han_ji}\t{tai_lo_im_piau}\t{weight}\t{summary if summary else ''}\t{create_time}\n")

        print(f"✅ RIME 字典 `{dict_filename}` 匯出完成！")
        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 匯出 RIME 字典失敗: {e}")
        return EXIT_CODE_FAILURE

    finally:
        conn.close()

# =========================================================================
# 主程式執行
# =========================================================================
def main():
    if len(sys.argv) > 1:
        mode = sys.argv[1]
    else:
        mode = "3"

    if mode == "3":
        return export_to_rime_dictionary()
    else:
        print("❌ 錯誤：請輸入有效模式 (3)")
        return EXIT_CODE_INVALID_INPUT

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)