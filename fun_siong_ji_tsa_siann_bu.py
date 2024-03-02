import xlwings as xw

def siong_ji_tsa_siann_bu(siong_ji):
    try:
        # 使用 xlwings 打開 Excel 檔案
        file_path = r'.\\tools\\反切上字與聲母對映表.xlsx'
        wb = xw.Book(file_path)
        sheet = wb.sheets['反切上字表']  # 假設工作表的名稱為 "反切上字"
        
        # 查找 siong_ji 在 "反切上字" 欄 (G2:G39) 的位置
        for row in range(2, 40):  # 假設資料從第2行開始，到第39行
            cell_value = sheet.range(f'G{row}').value
            if cell_value:
                words = cell_value.split()  # 按空格拆分儲存格的值
                if siong_ji in words:
                    # 如果找到了 siong_ji，則記錄下其相對映之 "列號"，並從相關欄位獲取資料
                    data = {
                        "id": row - 1,  # 第一行為標題行，故實際列號需要減1
                        "lui": sheet.range(f'B{row}').value,
                        "siann_bu": sheet.range(f'C{row}').value,
                        "tai_lo": sheet.range(f'D{row}').value,
                        "IPA": sheet.range(f'E{row}').value,
                    }
                    wb.close()  # 處理完成後關閉工作簿
                    return data
        wb.close()  # 如果沒找到，也關閉工作簿
        return None  # 如果在範圍內沒有找到 siong_ji，返回 None
    except Exception as e:
        print(f"發生錯誤：{e}")
        return None

# 測試函數
siong_ji = "普"  # 假設的反切上字
result = siong_ji_tsa_siann_bu(siong_ji)
print(result)
