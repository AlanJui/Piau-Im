"""
測試修改後的自動製作打字練習表功能
"""

import xlwings as xw


def test_char_10_handling():
    """
    測試 CHAR(10) 和換行符號的處理
    """
    print("=== 測試換行符號處理 ===")
    
    test_cases = [
        ('\n', True),      # 換行符號
        ('', True),        # 空字串  
        (10, True),        # CHAR(10) 數值
        ('φ', False),      # 終結符號（應該停止）
        ('漢', False),     # 正常漢字
        (None, True),      # None 值
    ]
    
    for i, (test_char, should_skip) in enumerate(test_cases, 1):
        # 模擬檢查邏輯
        skip = False
        
        # 檢查終結符號
        if test_char == 'φ':
            print(f"測試 {i}: {repr(test_char)} → 終結符號，停止處理")
            continue
            
        # 檢查換行符號
        if test_char == '\n' or str(test_char).strip() == '' or test_char == 10:
            skip = True
            
        # 檢查 None
        if test_char is None:
            skip = True
            
        result = "跳過" if skip else "處理"
        expected = "跳過" if should_skip else "處理"
        status = "✓" if (skip == should_skip) else "✗"
        
        print(f"測試 {i}: {repr(test_char):8s} → {result} {status} (期望: {expected})")


def check_excel_connection():
    """
    檢查 Excel 連接和工作表狀態
    """
    print("\n=== 檢查 Excel 連接狀態 ===")
    
    try:
        # 檢查作用中活頁簿
        wb = xw.books.active
        print(f"✓ 找到作用中活頁簿: {wb.name}")
        
        # 列出所有工作表
        sheet_names = [sheet.name for sheet in wb.sheets]
        print(f"✓ 工作表列表: {sheet_names}")
        
        # 檢查【漢字注音】工作表
        if '漢字注音' in sheet_names:
            han_ji_sheet = wb.sheets['漢字注音']
            print("✓ 找到【漢字注音】工作表")
            
            # 檢查一些範例資料
            print("\n範例資料檢查:")
            for col_index in range(4, 8):  # 檢查 D~G 欄
                col_letter = chr(64 + col_index)
                han_zi = han_ji_sheet.range(f'{col_letter}5').value
                pronunciation = han_ji_sheet.range(f'{col_letter}6').value
                print(f"  {col_letter}欄: 漢字={repr(han_zi)}, 標音={repr(pronunciation)}")
        else:
            print("❌ 找不到【漢字注音】工作表")
            
        # 檢查【打字練習表】工作表
        if '打字練習表' in sheet_names:
            print("✓ 找到【打字練習表】工作表")
        else:
            print("ℹ 【打字練習表】工作表不存在（程式執行時會建立）")
            
    except Exception as e:
        print(f"❌ Excel 連接錯誤: {e}")
        print("請確認:")
        print("- Excel 已開啟")
        print("- 有作用中的活頁簿")
        print("- 活頁簿中包含【漢字注音】工作表")


def main():
    """
    主測試函數
    """
    print("=== 自動製作打字練習表 - 修改功能測試 ===")
    print()
    
    # 測試換行符號處理邏輯
    test_char_10_handling()
    
    # 檢查 Excel 連接狀態
    check_excel_connection()
    
    print("\n=== 測試完成 ===")
    print("如果所有檢查都通過，可以執行主程式:")
    print("python auto_typing_practice.py")


if __name__ == "__main__":
    main()