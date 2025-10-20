"""
測試修正後的自動製作打字練習表功能
驗證多列處理和格式保持
"""

import xlwings as xw


def test_multi_row_logic():
    """
    測試多列處理邏輯
    """
    print("=== 測試多列處理邏輯 ===")
    
    # 計算各列的起始行號：3, 7, 11, 15, 19, 23
    row_starts = [3 + i * 4 for i in range(6)]
    print(f"列群組起始行: {row_starts}")
    
    for row_group_index, base_row in enumerate(row_starts):
        han_zi_row = base_row + 2    # 第3格
        pronunciation_row = base_row + 3  # 第4格
        
        print(f"第 {row_group_index + 1} 列群組:")
        print(f"  基準行: {base_row}")  
        print(f"  漢字行: {han_zi_row}")
        print(f"  標音行: {pronunciation_row}")
        
        # 示例欄位
        for col_index in [4, 5, 6]:  # D, E, F 欄
            col_letter = chr(64 + col_index)
            print(f"  {col_letter}欄: {col_letter}{han_zi_row}/{col_letter}{pronunciation_row}")
        print()


def test_excel_data_reading():
    """
    測試 Excel 資料讀取
    """
    print("=== 測試 Excel 資料讀取 ===")
    
    try:
        wb = xw.books.active
        han_ji_sheet = wb.sheets['漢字注音']
        print(f"✓ 成功連接到【漢字注音】工作表")
        
        # 測試讀取各列的資料
        row_starts = [3 + i * 4 for i in range(6)]
        
        for row_group_index, base_row in enumerate(row_starts):
            print(f"\n第 {row_group_index + 1} 列群組資料:")
            
            han_zi_row = base_row + 2
            pronunciation_row = base_row + 3
            
            # 檢查前幾欄的資料
            for col_index in range(4, 8):  # D~G 欄
                col_letter = chr(64 + col_index)
                try:
                    han_zi = han_ji_sheet.range(f'{col_letter}{han_zi_row}').value
                    pronunciation = han_ji_sheet.range(f'{col_letter}{pronunciation_row}').value
                    
                    print(f"  {col_letter}欄: 漢字={repr(han_zi)}, 標音={repr(pronunciation)}")
                    
                    # 如果遇到終結符號就停止
                    if han_zi == 'φ':
                        print(f"  *** 在 {col_letter}{han_zi_row} 遇到終結符號 ***")
                        return
                        
                except Exception as e:
                    print(f"  {col_letter}欄: 讀取錯誤 - {e}")
                    
    except Exception as e:
        print(f"❌ Excel 連接錯誤: {e}")


def test_format_preservation():
    """
    測試格式保持功能
    """
    print("\n=== 測試格式保持功能 ===")
    
    try:
        wb = xw.books.active
        
        # 檢查是否有打字練習表
        if '打字練習表' in [sheet.name for sheet in wb.sheets]:
            typing_sheet = wb.sheets['打字練習表']
            print("✓ 找到【打字練習表】工作表")
            
            # 檢查 B4:M4 的格式設定
            first_row_range = typing_sheet.range('B4:M4')
            print(f"✓ 第一列範圍: {first_row_range.address}")
            
            # 模擬格式複製測試
            print("格式複製邏輯:")
            print("  source_range = typing_sheet.range('B4:M4')")
            print("  target_range = typing_sheet.range('B5:M5')")
            print("  source_range.api.Copy()")
            print("  target_range.api.PasteSpecial(-4122)")
            print("  wb.app.api.CutCopyMode = False")
            
        else:
            print("ℹ 【打字練習表】工作表不存在")
            
    except Exception as e:
        print(f"❌ 測試錯誤: {e}")


def main():
    """
    主測試函數
    """
    print("=== 修正後程式功能測試 ===\n")
    
    # 測試多列處理邏輯
    test_multi_row_logic()
    
    # 測試 Excel 資料讀取
    test_excel_data_reading()
    
    # 測試格式保持
    test_format_preservation()
    
    print("\n=== 測試完成 ===")
    print("如果所有測試都通過，可以執行修正後的主程式:")
    print("python auto_typing_practice.py")


if __name__ == "__main__":
    main()