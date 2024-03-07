import xlwings as xw
import os


def main():
    # ===============================================================================
    # 1. 啟動 Excel ，建立一個 Excel 檔案
    # ===============================================================================
    # 啟動 Excel 應用程式
    app = xw.App(visible=True)

    # 新建一個工作簿
    wb = app.books.add()
    sheet = wb.sheets[0]

    # 設定 A 欄的寛度為128
    sheet.range("A:A").column_width = 128

    # 設定 A 欄各儲存格對於文字的處理為「自動換行」
    sheet.range("A:A").wrap_text = True

    # ===============================================================================
    # 2. 將子目錄 ./docs 的檔案： env.xlsx ，其中名稱為：env 的工作表，複製到這個
    # ===============================================================================
    # 新建立的 Excel 檔案中指定 env.xlsx 檔案的路徑
    # source_file_path = os.path.join("./docs", "env.xlsx")
    source_file_path = os.path.join(".\\output", "env.xlsx")

    # 開啟 env.xlsx 檔案
    source_wb = xw.Book(source_file_path)

    # 在新工作簿中複製工作表內容
    source_wb.sheets[0].copy(after=sheet)

    # ===============================================================================
    # 3. 將新建立的檔案，以名稱：Piaum-Im.xlsx 存檔
    # ===============================================================================
    # 指定新 Excel 檔案的路徑和名稱
    new_file_path = os.path.join(".\\output", "Piau-Tsu-Im.xlsx")

    # 儲存新建立的工作簿
    wb.save(new_file_path)

    # 關閉工作簿和 Excel 應用程式
    source_wb.close()
    # wb.close()
    # app.quit()

    # 選擇 Piau-Im.xlsx 檔案的〖工作表1〗工作表
    sheet.select()
