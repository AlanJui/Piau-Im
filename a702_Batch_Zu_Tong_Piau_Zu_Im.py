import os
import subprocess
import sys

import xlwings as xw

from p702_Ca_Han_Ji_Thak_Im import ca_han_ji_thak_im
from p710_thiam_han_ji import fill_hanji_in_cells
from p730_Tng_Sing_Bang_Iah import tng_sing_bang_iah

# 指定虛擬環境的 Python 路徑
venv_python = os.path.join(".venv", "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(".venv", "bin", "python")

# 定義檔案目錄
directory = r"C:\work\Piau-Im\output2"

# 所有檔案名稱
files = [
    "【河洛話注音】金剛般若波羅蜜經001。法會因由分第一.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經002。善現啟請分第二.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經003。大乘正宗分第三.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經004。妙行無住分第四.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經005。如理實見分第五.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經006。正信希有分第六.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經007。無得無說分第七.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經009。一相無相分第九.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經010。莊嚴淨土分第十.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經011。無為福勝分第十一.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經012。尊重正教分第十二.xlsx",
    # "【河洛話注音】金剛般若波羅蜜經013。如法受持分第十三.xlsx"
]

# 迴圈遍歷所有檔案並依次執行 Python 檔案
for file_name in files:
    file_path = os.path.join(directory, file_name)

    # 打開 Excel 檔案
    wb = xw.Book(file_path)

    # 顯示「已輸入之拼音字母及注音符號」 
    named_range = wb.names['顯示注音輸入']  # 選擇名為 "顯示注音輸入" 的命名範圍# 選擇名為 "顯示注音輸入" 的命名範圍
    named_range.refers_to_range.value = True

    # (1) A720: 將 V3 儲存格內的漢字，逐個填入標音用方格。
    # fill_hanji_in_cells(wb)     

    # (2) A731: 自動為漢字查找讀音，並抄寫到漢字的上方(拼音)及下方(注音)。
    ca_han_ji_thak_im(wb, '漢字注音', 'V3')

    # (3) A740: 將【漢字注音】工作表的內容，轉成 HTML 網頁檔案。
    tng_sing_bang_iah(wb, '漢字注音', 'V3')

    # (4) A750: 將 Tai_Gi_Zu_Im_Bun.xlsx 檔案，依 env 工作表的設定，另存新檔到指定目錄。
    setting_sheet = wb.sheets["env"]
    new_file_name = str(
        setting_sheet.range("C4").value
    ).strip()
    
    # 設定檔案輸出路徑，存於專案根目錄下的 output2 資料夾
    output_path = wb.names['OUTPUT_PATH'].refers_to_range.value 
    new_file_path = os.path.join(
        ".\\{0}".format(output_path), 
        f"【河洛話注音】{new_file_name}" + ".xlsx")

    # 儲存新建立的工作簿
    wb.save(new_file_path)

    # 保存 Excel 檔案
    wb.close()

