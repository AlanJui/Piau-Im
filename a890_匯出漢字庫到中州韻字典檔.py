# =========================================================================
# 載入程式所需套件/模組
# =========================================================================
import logging
import os
import shutil
import sqlite3
import sys
from datetime import datetime

from dotenv import load_dotenv

from mod_標音 import convert_tl_to_tlpa  # 載入台羅音標轉 TLPA 的函式

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
# 功能：將資料庫「漢字庫」資料表匯出為符合 RIME 字典格式的 YAML 檔
# =========================================================================
def export_database_to_rime_yaml():
    r"""
    從資料庫讀取【漢字庫】資料，產生符合中州韻輸入法引擎字典規格的 YAML 檔，
    檔名為 tl_ji_khoo_peh_ue.yaml，接著將此檔案複製到下列兩個目錄：
      - C:\Users\AlanJui\AppData\Roaming\Rime\
      - Z:\home\alanjui\workspace\rime\rime-tlpa\
    """
    conn = None
    try:
        # 連接資料庫並讀取資料表內容
        conn = sqlite3.connect(DB_HO_LOK_UE)
        cursor = conn.cursor()
        cursor.execute("SELECT 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間 FROM 漢字庫;")
        rows = cursor.fetchall()

        # ---------------------------------------------------------------------
        # 定義 RIME 字典檔的標頭內容（Header）
        # ---------------------------------------------------------------------
        header_content = """# Rime dictionary
# encoding: utf-8
#
# 河洛白話音
---
name: tl_ji_khoo_peh_ue
version: "0.1.0.0"
sort: by_weight
use_preset_vocabulary: false
columns:
  - text    # 漢字／詞彙
  - code    # 台灣音標（TLPA）拼音字母
  - weight  # 常用度（優先顯示度）
  - stem    # 用法舉例
  - create  # 建立日期
import_tables:
  - tl_ji_khoo_kah_kut_bun      # 甲骨文考證漢字庫
  # - tl_ji_khoo_peh_ue_cu_ting      # 個人自訂擴充字庫
  # - tl_ji_khoo_ciann_ji
  # - tl_ji_khoo_siong_iong_si_lui
  # - tl_ji_khoo_tai_uan_si_lui
...
"""
        # ---------------------------------------------------------------------
        # 處理資料表中每一筆資料，轉換成符合字典檔格式的資料行
        # 資料行以 tab 字元分隔各欄： text, code, weight, stem, create
        # ---------------------------------------------------------------------
        data_lines = []
        for row in rows:
            # 資料表欄位依序： 識別號, 漢字, 台羅音標, 常用度, 摘要說明, 更新時間
            text = row[1] if row[1] is not None else ""
            # 將「台羅音標」轉換為 TLPA
            code = convert_tl_to_tlpa(row[2]) if row[2] is not None else ""
            weight = str(row[3]) if row[3] is not None else ""
            stem = row[4] if row[4] is not None else ""
            create = row[5] if row[5] is not None else ""
            # 組成一行（以 tab 分隔）
            line = f"{text}\t{code}\t{weight}\t{stem}\t{create}"
            data_lines.append(line)

        # ---------------------------------------------------------------------
        # 將 header 與資料行寫入檔案
        # ---------------------------------------------------------------------
        output_filename = "tl_ji_khoo_peh_ue.dict.yaml"
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(header_content)
            f.write("\n")  # 確保 header 與資料間有換行
            f.write("\n".join(data_lines))

        print("✅ RIME 字典檔已產生:", output_filename)
        logging.info("RIME dictionary exported: %s", output_filename)

        # ---------------------------------------------------------------------
        # 將產生的 YAML 檔案複製到指定的兩個目錄
        # ---------------------------------------------------------------------
        dest_dirs = [
            r"C:\Users\AlanJui\AppData\Roaming\Rime",
            r"Z:\home\alanjui\workspace\rime\rime-tlpa"
        ]
        for dest in dest_dirs:
            if not os.path.exists(dest):
                logging.warning("目標目錄不存在: %s", dest)
                print(f"⚠️ 目標目錄不存在: {dest}")
                continue
            dest_file = os.path.join(dest, output_filename)
            shutil.copy(output_filename, dest_file)
            print(f"✅ RIME 字典檔已複製到 {dest}")
            logging.info("RIME dictionary copied to: %s", dest_file)

        return EXIT_CODE_SUCCESS

    except Exception as e:
        print(f"❌ 匯出 RIME 字典檔失敗: {e}")
        logging.error("Error exporting RIME dictionary: %s", e)
        return EXIT_CODE_FAILURE

    finally:
        if conn:
            conn.close()


# =========================================================================
# 主程式執行
# =========================================================================
def main():
    """
    執行模式預設為 rime（匯出 RIME 字典檔），
    也可從命令列傳入參數來指定模式，例如：
      python this_script.py rime
    """
    mode = sys.argv[1] if len(sys.argv) > 1 else "rime"

    if mode == "rime":
        return export_database_to_rime_yaml()
    else:
        print("❌ 錯誤：請輸入有效模式 ('rime')")
        return EXIT_CODE_INVALID_INPUT


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
