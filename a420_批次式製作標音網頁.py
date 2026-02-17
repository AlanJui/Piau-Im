import os
import sys

import xlwings as xw

from mod_file_access import ensure_extension_name
from p730_Tng_Sing_Bang_Iah_R1 import tng_sing_bang_iah

# 指定虛擬環境的 Python 路徑
venv_python = os.path.join(".venv", "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(".venv", "bin", "python")

# 定義檔案目錄
directory = r"C:\work\Piau-Im\output2"

# 所有檔案名稱
files = [
    '【河洛話注音】金剛般若波羅蜜經001',
    '【河洛話注音】金剛般若波羅蜜經002',
    '【河洛話注音】金剛般若波羅蜜經003',
    '【河洛話注音】金剛般若波羅蜜經004',
    '【河洛話注音】金剛般若波羅蜜經005',
    '【河洛話注音】金剛般若波羅蜜經006',
    '【河洛話注音】金剛般若波羅蜜經007',
    '【河洛話注音】金剛般若波羅蜜經008',
    '【河洛話注音】金剛般若波羅蜜經009',
    '【河洛話注音】金剛般若波羅蜜經010',
    '【河洛話注音】金剛般若波羅蜜經011',
    '【河洛話注音】金剛般若波羅蜜經012',
    '【河洛話注音】金剛般若波羅蜜經013',
    '【河洛話注音】金剛般若波羅蜜經014',
    '【河洛話注音】金剛般若波羅蜜經015',
    '【河洛話注音】金剛般若波羅蜜經016',
    '【河洛話注音】金剛般若波羅蜜經017',
    '【河洛話注音】金剛般若波羅蜜經018',
    '【河洛話注音】金剛般若波羅蜜經019',
    '【河洛話注音】金剛般若波羅蜜經020',
    '【河洛話注音】金剛般若波羅蜜經021',
    '【河洛話注音】金剛般若波羅蜜經022',
    '【河洛話注音】金剛般若波羅蜜經023',
    '【河洛話注音】金剛般若波羅蜜經024',
    '【河洛話注音】金剛般若波羅蜜經025',
    '【河洛話注音】金剛般若波羅蜜經026',
    '【河洛話注音】金剛般若波羅蜜經027',
    '【河洛話注音】金剛般若波羅蜜經028',
    '【河洛話注音】金剛般若波羅蜜經029',
    '【河洛話注音】金剛般若波羅蜜經030',
    '【河洛話注音】金剛般若波羅蜜經031',
    '【河洛話注音】金剛般若波羅蜜經032',
]

# 迴圈遍歷所有檔案並依次執行 Python 檔案
for file_name in files:
    updated_file_name = ensure_extension_name(file_name, 'xlsx')
    file_path = os.path.join(directory, updated_file_name)

    # 打開 Excel 檔案
    wb = xw.Book(file_path)

    # 顯示「已輸入之拼音字母及注音符號」
    named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
    named_range.refers_to_range.value = True

    # (1) A740: 將【漢字注音】工作表的內容，轉成 HTML 網頁檔案。
    tng_sing_bang_iah(wb, '漢字注音', 'V3', '去頁頭')

    # (2) A750: 將 Tai_Gi_Zu_Im_Bun.xlsx 檔案，依 env 工作表的設定，另存新檔到指定目錄。
    try:
        file_name = str(wb.names['TITLE'].refers_to_range.value).strip()
    except KeyError:
        # print("未找到命名範圍 'TITLE'，使用預設名稱")
        # file_name = "default_file_name.xlsx"  # 提供一個預設檔案名稱
        setting_sheet = wb.sheets["env"]
        file_name = str(
            setting_sheet.range("C4").value
        ).strip()

    # 設定檔案輸出路徑，存於專案根目錄下的 output2 資料夾
    output_path = wb.names['OUTPUT_PATH'].refers_to_range.value
    new_file_path = os.path.join(
        ".\\{0}".format(output_path),
        f"【河洛話注音】{file_name}.xlsx")

    # 儲存新建立的工作簿
    wb.save(new_file_path)

    # 保存 Excel 檔案
    wb.close()

