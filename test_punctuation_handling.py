"""
測試修正後的標點符號和換行處理功能
"""

import xlwings as xw

from a710_製作打字練習工作表 import is_line_break, is_punctuation


def test_punctuation_detection():
    """
    測試標點符號識別功能
    """
    print("=== 測試標點符號識別 ===")

    test_cases = [
        # 中文標點
        ('，', True),
        ('。', True),
        ('！', True),
        ('？', True),
        ('；', True),
        ('：', True),
        ('「', True),
        ('」', True),
        ('《', True),
        ('》', True),

        # 英文標點
        (',', True),
        ('.', True),
        ('!', True),
        ('?', True),
        (';', True),
        (':', True),
        ('"', True),
        ("'", True),
        ('(', True),
        (')', True),

        # 非標點符號
        ('漢', False),
        ('字', False),
        ('A', False),
        ('1', False),
        ('ㄅ', False),
        (None, False),
        ('', False),
    ]

    for i, (char, expected) in enumerate(test_cases, 1):
        result = is_punctuation(char)
        status = "✓" if result == expected else "✗"
        print(f"測試 {i:2d}: {repr(char):8s} → {result} {status} (期望: {expected})")


def test_line_break_detection():
    """
    測試換行控制字元識別功能
    """
    print("\n=== 測試換行控制字元識別 ===")

    test_cases = [
        ('\n', True),      # 換行符號
        ('', True),        # 空字串
        ('   ', True),     # 空白字串
        (10, True),        # CHAR(10)
        (None, False),     # None
        ('漢', False),     # 正常字元
        ('《', False),     # 標點符號
    ]

    for i, (char, expected) in enumerate(test_cases, 1):
        result = is_line_break(char)
        status = "✓" if result == expected else "✗"
        print(f"測試 {i:2d}: {repr(char):8s} → {result} {status} (期望: {expected})")


def test_template_sheet_access():
    """
    測試模版工作表訪問
    """
    print("\n=== 測試模版工作表訪問 ===")

    try:
        wb = xw.books.active
        print(f"✓ 成功連接活頁簿: {wb.name}")

        sheet_names = [sheet.name for sheet in wb.sheets]
        print(f"✓ 所有工作表: {sheet_names}")

        # 檢查【打字練習表（模版）】工作表
        template_sheet_names = ['打字練習表（模版）', '打字練習表 (模版)']
        template_found = False

        for template_name in template_sheet_names:
            if template_name in sheet_names:
                template_sheet = wb.sheets[template_name]
                print(f"✓ 找到【{template_name}】工作表")

                # 檢查模版範圍
                template_range = template_sheet.range('B4:M4')
                print(f"✓ 模版範圍: {template_range.address}")

                # 測試格式複製邏輯
                print("格式複製邏輯測試:")
                print("  template_range.api.Copy()")
                print("  target_range.api.PasteSpecial(-4122)")
                print("  wb.app.api.CutCopyMode = False")

                template_found = True
                break

        if not template_found:
            print("❌ 找不到模版工作表")
            print("請確認是否有以下任一工作表:")
            for name in template_sheet_names:
                print(f"  - {name}")

    except Exception as e:
        print(f"❌ 測試錯誤: {e}")


def test_processing_logic():
    """
    測試處理邏輯
    """
    print("\n=== 測試處理邏輯 ===")

    test_data = [
        ('漢', 'ㄏㄢˋ', '正常漢字處理'),
        ('《', None, '標點符號處理（只填B欄）'),
        ('\n', None, '換行控制字元（留空白行）'),
        ('字', 'ㄗˋ', '正常漢字處理'),
        ('。', None, '標點符號處理（只填B欄）'),
        ('φ', None, '終結符號（停止處理）'),
    ]

    print("模擬處理流程:")
    current_row = 4

    for han_zi, pronunciation, description in test_data:
        print(f"\n第 {current_row} 列:")
        print(f"  輸入: 漢字={repr(han_zi)}, 標音={repr(pronunciation)}")

        if han_zi == 'φ':
            print(f"  處理: {description}")
            break
        elif is_line_break(han_zi):
            print(f"  處理: {description}")
            current_row += 1
        elif is_punctuation(han_zi):
            print(f"  處理: {description}")
            print(f"  結果: B{current_row}='{han_zi}', C{current_row}=空白")
            current_row += 1
        else:
            print(f"  處理: {description}")
            print(f"  結果: B{current_row}='{han_zi}', C{current_row}='{pronunciation}', E{current_row}~M{current_row}=分解字元")
            current_row += 1


def main():
    """
    主測試函數
    """
    print("=== 修正功能測試 ===")
    print("測試標點符號和換行處理功能\n")

    # 測試標點符號識別
    test_punctuation_detection()

    # 測試換行控制字元識別
    test_line_break_detection()

    # 測試模版工作表訪問
    test_template_sheet_access()

    # 測試處理邏輯
    test_processing_logic()

    print("\n=== 測試完成 ===")
    print("如果所有測試都通過，可以執行修正後的主程式:")
    print("python auto_typing_practice.py")


if __name__ == "__main__":
    main()