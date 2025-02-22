# =========================================================================
# 載入程式所需套件/模組/函式庫
# =========================================================================
import asyncio
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import aiohttp
import requests
import xlwings as xw
from dotenv import load_dotenv

# 載入自訂模組/函式
from mod_excel_access import (
    convert_to_excel_address,
    ensure_sheet_exists,
    excel_address_to_row_col,
    get_value_by_name,
)
from mod_標音 import convert_tlpa_to_tl

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤

# =========================================================================
# 載入環境變數
# =========================================================================
load_dotenv()

# 預設檔案名稱從環境變數讀取
DB_HO_LOK_UE = os.getenv('DB_HO_LOK_UE', 'Ho_Lok_Ue.db')
DB_KONG_UN = os.getenv('DB_KONG_UN', 'Kong_Un.db')

# =========================================================================
# 設定日誌
# =========================================================================
from mod_logging import init_logging, logging_exc_error, logging_process_step

init_logging()

# =========================================================================
# 程式區域函式
# =========================================================================
# 聲調符號對應表（帶調號母音 → 對應數字）
tone_mapping = {
    "a̍": ("a", "8"), "á": ("a", "2"), "ǎ": ("a", "6"), "â": ("a", "5"), "ā": ("a", "7"), "à": ("a", "3"),
    "e̍": ("e", "8"), "é": ("e2"), "ě": ("e6"), "ê": ("e5"), "ē": ("e7"), "è": ("e3"),
    "i̍": ("i", "8"), "í": ("i2"), "ǐ": ("i6"), "î": ("i5"), "ī": ("i7"), "ì": ("i3"),
    "o̍": ("o", "8"), "ó": ("o2"), "ǒ": ("o6"), "ô": ("o5"), "ō": ("o7"), "ò": ("o3"),
    "u̍": ("u", "8"), "ú": ("u2"), "ǔ": ("u6"), "û": ("u5"), "ū": ("u7"), "ù": ("u3"),
    "m̍": ("m", "8"), "ḿ": ("m2"), "m̀": ("m3"), "m̂": ("m5"), "m̄": ("m7"),
    "n̍": ("n", "8"), "ń": ("n2"), "ň": ("n6"), "n̂": ("n5"), "n̄": ("n7")
}

# 聲母轉換規則（台羅拼音 → 台語音標+）
initials_mapping = {
    "tsh": "c",
    "ts": "z"
}

def convert_tai_lo_to_tlpa_plus(im_piau: str) -> str:
    """
    先將台羅拼音的聲調符號轉換為 TLPA 數字聲調，然後轉換聲母為台語音標+（TLPA+）。
    例如：
        - lio̍k → liok8
        - tāi → tai7
        - bô → bo5
        - siâu → siau5
        - tshit → cit8
        - tsuan → zuan5
    """
    tone_number = ""

    # 1. 先處理聲調轉換
    for tone_mark, (base_char, number) in tone_mapping.items():
        if tone_mark in im_piau:
            im_piau = im_piau.replace(tone_mark, base_char)  # 移除調號，還原原始母音
            tone_number = number  # 記錄對應的聲調數字
            break  # 只會有一個聲調符號，找到就停止

    # 2. 將聲母轉換為 TLPA+
    for tai_lo, tlpa in initials_mapping.items():
        if im_piau.startswith(tai_lo):  # 只轉換開頭的聲母部分
            im_piau = tlpa + im_piau[len(tai_lo):]
            break

    # 3. 若有聲調數字，則加到末尾
    if tone_number:
        return im_piau + tone_number

    return im_piau  # 若無聲調符號則不變更

def query_tai_gi_han_ji(han_ji: str):
    """
    查詢臺灣閩南語常用詞辭典 API，獲取漢字的台語音標與字義。

    參數：
        han_ji (str): 要查詢的漢字

    回傳：
        tok_im (str): 台語音標（TLPA 拼音），查無讀音則回傳 "N/A"
        explanations (str): 字義（多個解釋合併為字串），查無解釋則回傳 "N/A"
    """
    url = f"https://www.moedict.tw/t/{han_ji}.json"
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()
        if "h" in data:
            tok_im = data["h"][0].get("T", "N/A")  # 讀音
            explanations = "；".join([item["f"].replace("`", "") for item in data["h"][0].get("d", [])]) or "N/A"  # 字義
            return tok_im, explanations

    return "N/A", "N/A"

def update_excel_with_tai_gi(wb):
    """
    讀取 Excel 檔案，為 A 欄的每個漢字查詢台語音標與字義，並填入 C 欄（台語音標）與 F 欄（字義）。

    參數：
        file_path (str): Excel 檔案的路徑
    """
    try:
        sheet = wb.sheets["缺字表2"]  # 選擇工作表
    except Exception as e:
        logging_exc_error(f"找不到名為「缺字表2」的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    row = 2  # 從第 2 列開始（跳過標題列）
    while True:
        han_ji = sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束
            break

        print(f"({row - 1}) 漢字：{han_ji} ...")

        # 查詢台語音標與字義
        tok_im, explanations = query_tai_gi_han_ji(han_ji)

        # 更新 Excel 的 C 欄（台語音標）與 F 欄（字義）
        tai_gi_im_piau = convert_tai_lo_to_tlpa_plus(tok_im)  # 將台語音標轉換為 TLPA+
        sheet.range(f"C{row}").value = tai_gi_im_piau
        sheet.range(f"F{row}").value = explanations

        print(f"台語音標：{tok_im}, 字義：{explanations}")

        time.sleep(0.5)  # 避免 API 請求過快
        row += 1  # 讀取下一行

    return EXIT_CODE_SUCCESS

async def query_tai_gi_han_ji_async(han_ji: str):
    url = f"https://www.moedict.tw/t/{han_ji}.json"
    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            if response.status == 200:
                data = await response.json()
                if "h" in data:
                    tok_im = data["h"][0].get("T", "N/A")  # 讀音
                    explanations = "；".join([item["f"].replace("`", "") for item in data["h"][0].get("d", [])]) or "N/A"  # 字義
                    return tok_im, explanations
    return "N/A", "N/A"

async def update_excel_with_tai_gi_async(wb):
    try:
        sheet = wb.sheets["缺字表2"]  # 選擇工作表
    except Exception as e:
        logging_exc_error(f"找不到名為「缺字表2」的工作表", e)
        return EXIT_CODE_INVALID_INPUT

    row = 2  # 從第 2 列開始（跳過標題列）
    while True:
        han_ji = sheet.range(f"A{row}").value  # 讀取 A 欄（漢字）
        if not han_ji:  # 若 A 欄為空，則結束
            break

        print(f"({row - 1}) 漢字：{han_ji} ...")

        # 查詢台語音標與字義
        tok_im, explanations = await query_tai_gi_han_ji_async(han_ji)

        # 更新 Excel 的 C 欄（台語音標）與 F 欄（字義）
        tai_gi_im_piau = convert_tai_lo_to_tlpa_plus(tok_im)  # 將台語音標轉換為 TLPA+
        sheet.range(f"C{row}").value = tai_gi_im_piau
        sheet.range(f"F{row}").value = explanations

        print(f"台語音標：{tok_im}, 字義：{explanations}")

        row += 1  # 讀取下一行

    return EXIT_CODE_SUCCESS

async def main_async():
    # =========================================================================
    # 初始化
    # =========================================================================
    # 取得專案根目錄。
    current_file_path = Path(__file__).resolve()
    project_root = current_file_path.parent
    # 取得程式名稱
    # program_file_name = current_file_path.name
    program_name = current_file_path.stem

    logging_process_step(f"《========== 程式開始執行：{program_name} ==========》")
    logging_process_step(f"專案根目錄為: {project_root}")

    try:
        wb = xw.apps.active.books.active
    except Exception as e:
        logging_exc_error(f"找不到作用中活頁簿檔", e)
        return EXIT_CODE_INVALID_INPUT

    status_code = await update_excel_with_tai_gi_async(wb)
    if status_code != EXIT_CODE_SUCCESS:
        logging_process_step(f"程式執行失敗，錯誤代碼：{status_code}")
        return status_code

    return EXIT_CODE_SUCCESS

if __name__ == "__main__":
    exit_code = asyncio.run(main_async())
    sys.exit(exit_code)