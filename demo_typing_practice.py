"""
簡單的使用範例
展示如何使用自動製作打字練習表程式
"""

import xlwings as xw

from a710_製作拚音打字練習工作表 import create_typing_practice_sheet


def demo_usage():
    """
    展示使用方法
    """
    print("=== 自動製作打字練習表 - 使用範例 ===")
    print()
    print("1. 請確保 Excel 已開啟，且包含【漢字注音】工作表")
    print("2. 確認【漢字注音】工作表的格式正確")
    print("3. 準備執行程式...")
    print()

    try:
        # 檢查是否有作用中的 Excel 活頁簿
        wb = xw.books.active
        print(f"✓ 找到作用中活頁簿: {wb.name}")

        # 檢查是否有【漢字注音】工作表
        if '漢字注音' in [sheet.name for sheet in wb.sheets]:
            print("✓ 找到【漢字注音】工作表")

            # 執行製作打字練習表
            print("\n開始製作打字練習表...")
            success = create_typing_practice_sheet()

            if success:
                print("\n🎉 打字練習表製作完成！")
                print("請查看【打字練習表】工作表")
            else:
                print("\n❌ 打字練習表製作失敗")
        else:
            print("❌ 找不到【漢字注音】工作表")
            print("   可用的工作表:", [sheet.name for sheet in wb.sheets])

    except Exception as e:
        print(f"❌ 發生錯誤: {e}")
        print("\n請確認:")
        print("- Excel 已開啟")
        print("- 有作用中的活頁簿")
        print("- 活頁簿中包含【漢字注音】工作表")


if __name__ == "__main__":
    demo_usage()