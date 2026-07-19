# =========================================================================
# a820_匯出製成中州韻字典檔【台語注音二式】.py
#
# 功能說明：
# 將【Ho_Lok_Ue.db】/【漢字庫】資料表內的漢字讀音紀錄匯出，製成
# 中州韻輸入方案字典檔（.yaml）檔案：ji_khoo_bpm2.dict.yaml。
#
# 此字典檔專供「台語注音二式」輸入方案使用，其羅馬拼音系統採用「台語注音二式」。
#
# 若需要詳細之「台語注音二式」羅馬拼音系統說明，可參考：
# C:\Users\AlanJui\work\rime-tlpa\docs\090_漢字標音轉換指引.md 文件。
# =========================================================================


# =========================================================================
# 載入程式所需套件/模組
# =========================================================================
import logging
import os
import shutil
import sqlite3
import sys

from dotenv import load_dotenv

from mod_convert_TLPA_to_MPS2 import convert_TLPA_to_MPS2  # 載入 TLPA 轉台語注音二式的函式
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
    檔名為 ji_khoo_bpm2.dict.yaml；code 欄之拼音採【台語注音二式（BPM2）】，
    轉換路徑：台羅音標（資料庫存放格式）→ TLPA → 台語注音二式。
    接著將此檔案複製到下列兩個目錄：
      - C:\Users\AlanJui\AppData\Roaming\Rime\
      - C:\Users\AlanJui\work\rime-tlpa\
    """
    conn = None
    try:
        # 連接資料庫並讀取資料表內容
        # 【漢字庫】資料表現行結構：識別號、漢字、台羅音標、常用度、摘要說明、更新時間、最近揀用時間。
        # 排序規則與查音邏輯（mod_ca_ji_tian.py）一致：同漢字之多筆讀音，
        # 依【常用度】由大至小；常用度相同時，依【最近揀用時間】由新至舊。
        conn = sqlite3.connect(DB_HO_LOK_UE)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT 漢字, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間
            FROM 漢字庫
            ORDER BY 漢字 ASC,
                     COALESCE(常用度, 0) DESC,
                     COALESCE(最近揀用時間, '') DESC;
            """
        )
        rows = cursor.fetchall()

        # ---------------------------------------------------------------------
        # 定義 RIME 字典檔的標頭內容（Header）
        # ---------------------------------------------------------------------
        header_content = """# Rime dictionary
# encoding: utf-8
#
# Ho_Lok_Ue.db/漢字庫資料表轉製成中州韻輸入方案字典檔
# 此字典檔專供「台語注音二式」輸入方案使用，其羅馬拼音系統採用「台語注音二式」。
---
name: ji_khoo_bpm2
version: "v0.1.0"
sort: by_weight
use_preset_vocabulary: false
columns:
  - text    # 漢字
  - code    # 台語注音二式（BPM2）拼音
  - weight  # 常用度（優先顯示度）
  - stem    # 用法舉例
  - create  # 建立時間
import_tables:
  - ji_khoo_ziann_ji_bpm2
  - ji_khoo_ban_lam_bpm2
  - ji_khoo_su_lui_bpm2
...
"""
        # ---------------------------------------------------------------------
        # 處理資料表中每一筆資料，轉換成符合字典檔格式的資料行
        # 資料行以 tab 字元分隔各欄： text, code, weight, stem, create
        # ---------------------------------------------------------------------
        data_lines = []
        for row in rows:
            # 查詢結果欄位依序：漢字, 台羅音標, 常用度, 摘要說明, 更新時間, 最近揀用時間
            han_ji, tai_lo_im_piau, siong_iong_too, zik_iau, kenn_sin_si, kin_king_si = row
            text = han_ji if han_ji is not None else ""
            # 將「台羅音標」轉換為 TLPA，再由 TLPA 轉換為「台語注音二式」
            if tai_lo_im_piau is not None:
                tlpa_im_piau = convert_tl_to_tlpa(str(tai_lo_im_piau).strip().lower()) or ""
                code = convert_TLPA_to_MPS2(tlpa_im_piau)
            else:
                code = ""
            weight = str(siong_iong_too) if siong_iong_too is not None else ""
            # 摘要說明若為 NULL 或空白，一律填入 'NA' 佔位：避免該欄留空時，
            # 在無法顯示 Tab 控制字元的文字編輯器中，被誤認為多餘空白而遭刪除，
            # 破壞 RIME 字典檔以 Tab 分欄的結構。
            stem = zik_iau if zik_iau is not None and str(zik_iau).strip() != "" else "NA"
            # create 欄：取【最近揀用時間】（人工揀用之讀音較具參考性），無則取【更新時間】
            create = kin_king_si or kenn_sin_si or ""
            # 組成一行（以 tab 分隔）
            line = f"{text}\t{code}\t{weight}\t{stem}\t{create}"
            data_lines.append(line)

        # ---------------------------------------------------------------------
        # 將 header 與資料行寫入檔案
        # ---------------------------------------------------------------------
        output_filename = "ji_khoo_bpm2.dict.yaml"
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
            r"C:\Users\AlanJui\work\rime-tlpa"
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
