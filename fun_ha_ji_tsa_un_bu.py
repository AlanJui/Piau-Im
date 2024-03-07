import xlwings as xw

def ha_ji_tsa_un_bu(ha_ji):
    try:
        # 使用 xlwings 打開 Excel 檔案
        file_path = r'.\\tools\\反切下字與韻母對映表.xlsx'
        wb = xw.Book(file_path)
        sheet = wb.sheets['反切下字表']  
        
        # 查找 ha_ji 在指定範圍 (J2:J187) 的位置
        for row in range(2, 188):  # 資料從第2列開始，到第187列
            cell_value = sheet.range(f'J{row}').value
            if cell_value:
                words = cell_value.split()  # 按空格拆分儲存格的值
                if ha_ji in words:
                    # 如果找到了 ha_ji，則記錄下其相對映之 "列號"，並從相關欄位獲取資料
                    data = {
                        "id": row - 1,  # 第一行為標題行，故實際列號需要減1
                        "liap": sheet.range(f'B{row}').value, # 攝
                        "un_he": sheet.range(f'C{row}').value,  # 韻系
                        "si_siann": sheet.range(f'D{row}').value,  # 四聲 (平上去入)
                        "ting_de": sheet.range(f'E{row}').value,  # 等第
                        "khai_hap": sheet.range(f'F{row}').value,  # 開合
                        "un_bu": sheet.range(f'G{row}').value,  # 韻母
                        "tai_lo": sheet.range(f'H{row}').value,  # 臺羅拼音
                        "IPA": sheet.range(f'I{row}').value,  # 國際音標
                    }
                    wb.close()  # 處理完成後關閉工作簿
                    return data
        wb.close()  # 如果沒找到，也關閉工作簿
        return None  # 如果在範圍內沒有找到 ha_ji，返回 None
    except Exception as e:
        print(f"發生錯誤：{e}")
        return None


# 測試函數
ha_ji = "荅"  # 假設的反切上字
result = ha_ji_tsa_un_bu(ha_ji)
print(result)