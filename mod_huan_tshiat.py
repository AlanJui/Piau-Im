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
                        "tshing_lo": sheet.range(f'F{row}').value,
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
    

def e_ji_tsa_un_bu(e_ji):
    try:
        # 使用 xlwings 打開 Excel 檔案
        file_path = r'.\\tools\\反切下字與韻母對映表.xlsx'
        wb = xw.Book(file_path)
        sheet = wb.sheets['反切下字表']  
        
        # 查找 e_ji 在指定範圍 (J2:J187) 的位置
        for row in range(2, 188):  # 資料從第2列開始，到第187列
            cell_value = sheet.range(f'J{row}').value
            if cell_value:
                words = cell_value.split()  # 按空格拆分儲存格的值
                if e_ji in words:
                    # 如果找到了 e_ji，則記錄下其相對映之 "列號"，並從相關欄位獲取資料
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
        return None  # 如果在範圍內沒有找到 e_ji，返回 None
    except Exception as e:
        print(f"發生錯誤：{e}")
        return None

if __name__ == "__main__":
    # 測試 siong_ji_tsa_siann_bu 函數
    siong_ji = "普"
    result_siong_ji = siong_ji_tsa_siann_bu(siong_ji)
    assert result_siong_ji["siann_bu"] == "滂"
    assert result_siong_ji["tai_lo"] == "ph"
    assert result_siong_ji["tshing_lo"] == "次清"
    print(f"\n查反切上字：{siong_ji}")
    print(f"siong_ji_tsa_siann_bu 測試結果：{result_siong_ji}")

    # 測試 e_ji_tsa_un_bu 函數
    e_ji = "荅"
    result_e_ji = e_ji_tsa_un_bu(e_ji)
    assert result_e_ji["un_bu"] == "合"
    assert result_e_ji["tai_lo"] == "ap"
    print(f"\n查反切下字：{e_ji}")
    print(f"e_ji_tsa_un_bu 測試結果：{result_e_ji}")
