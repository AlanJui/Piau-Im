# =========================================================================
# 當 Tai_Gi_Zu_Im_Bun.xlsx 檔案已完成人工手動注音後，執行此程式可完成以下工作：
# (1) A730: 將人工填入之拼音及注音，抄寫到漢字的上方(拼音)及下方(注音)。
# (2) A740: 將【漢字注音】工作表的內容，轉成 HTML 網頁檔案。
# (3) A750: 將 Tai_Gi_Zu_Im_Bun.xlsx 檔案，依 env 工作表的設定，另存新檔到指定目錄。
# =========================================================================
import os
import subprocess
import sys

# 指定虛擬環境的 Python 路徑
venv_python = os.path.join(".venv", "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(".venv", "bin", "python")

# 依次執行三個 Python 檔案

# (1) A720: 將 V3 儲存格內的漢字，填入標音用方格。
subprocess.run([venv_python, "a720_待注音漢字填入標音用方格.py", "-i", "Tai_Gi_Zu_Im_Bun.xlsx"])

# (2) A731: 自動為漢字查找讀音，並抄寫到漢字的上方(拼音)及下方(注音)。
subprocess.run([venv_python, "a731_自動為漢字查注音.py", "-i", "Tai_Gi_Zu_Im_Bun.xlsx"])

# (3) A740: 將【漢字注音】工作表的內容，轉成 HTML 網頁檔案。
subprocess.run([venv_python, "a740_漢字注音轉網頁.py", "-i", "Tai_Gi_Zu_Im_Bun.xlsx"])

# (4) A750: 將 Tai_Gi_Zu_Im_Bun.xlsx 檔案，依 env 工作表的設定，另存新檔到指定目錄。
# subprocess.run([venv_python, "a750_漢字注音存檔.py"])
subprocess.run([venv_python, "a750_漢字注音存檔.py", "-i", "Tai_Gi_Zu_Im_Bun.xlsx"])
