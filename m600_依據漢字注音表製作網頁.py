import xlwings as xw

import settings
from p210_hoo_goa_chu_im_all import main_run as hoo_goa_chu_im_all

# ===========================================================================
# (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
# ===========================================================================
# 取得 Input 檔案名稱
file_path = settings.get_input_file_path()
if not file_path:
    print("未設定 .env 檔案")
    CONVERT_FILE_NAME = "hoo-goa-chu-im.xlsx"
else:
    CONVERT_FILE_NAME = file_path
print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

file_path = CONVERT_FILE_NAME
wb = xw.Book(file_path)


# ===========================================================================
# (1) 取得需要注音的「檔案名稱」及其「目錄路徑」。
# ===========================================================================
# 取得 Input 檔案名稱
file_path = settings.get_input_file_path()
if not file_path:
    print("未設定 .env 檔案")
    # sys.exit(2)
    opts = myfunc(sys.argv)
    if opts["input"] != "":
        CONVERT_FILE_NAME = opts["input"]
    else:
        CONVERT_FILE_NAME = "hoo-goa-chu-im.xlsx"
else:
    CONVERT_FILE_NAME = file_path
print(f"CONVERT_FILE_NAME = {CONVERT_FILE_NAME}")

# ===========================================================================
# (2) 將已注音之「漢字注音表」，製作成 HTML 格式之「注音／拼音／標音」網頁。
# ===========================================================================
hoo_goa_chu_im_all(CONVERT_FILE_NAME)
