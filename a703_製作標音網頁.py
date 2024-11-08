import os
import sys

import xlwings as xw

from mod_file_access import get_named_value
from p730_Tng_Sing_Bang_Iah import tng_sing_bang_iah

# 指定虛擬環境的 Python 路徑
venv_python = os.path.join(".venv", "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(".venv", "bin", "python")

# (0) 取得專案根目錄。
# 使用已打開且處於作用中的 Excel 工作簿
try:
    wb = xw.apps.active.books.active
except Exception as e:
    print(f"發生錯誤: {e}")
    print("無法找到作用中的 Excel 工作簿")
    sys.exit(2)

# 獲取活頁簿的完整檔案路徑
file_path = wb.fullname
print(f"完整檔案路徑: {file_path}")

# 獲取活頁簿的檔案名稱（不包括路徑）
file_name = wb.name
print(f"檔案名稱: {file_name}")

# 顯示「已輸入之拼音字母及注音符號」
# named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
# named_range.refers_to_range.value = True
named_range = get_named_value(wb, '顯示注音輸入', True)

# (1) A720: 將 V3 儲存格內的漢字，逐個填入標音用方格。
sheet = wb.sheets['漢字注音']   # 選擇工作表
sheet.activate()               # 將「漢字注音」工作表設為作用中工作表
sheet.range('A1').select()     # 將 A1 儲存格設為作用儲存格

# (2) A740: 將【漢字注音】工作表的內容，轉成 HTML 網頁檔案。
han_ji_piau_im_huat = get_named_value(wb, '標音方法')
tng_sing_bang_iah(wb=wb, sheet_name='漢字注音', cell='V3', page_type='含頁頭')

# (3) A750: 將 Tai_Gi_Zu_Im_Bun.xlsx 檔案，依 env 工作表的設定，另存新檔到指定目錄。
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
# wb.close()

