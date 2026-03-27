# import sys; sys.path.insert(0, ".")

import unittest

import xlwings as xw

from mod_字庫 import JiKhooDict

# =========================================================================
# 常數定義
# =========================================================================
# 定義 Exit Code
EXIT_CODE_SUCCESS = 0  # 成功
EXIT_CODE_NO_FILE = 1  # 無法找到檔案
EXIT_CODE_INVALID_INPUT = 2  # 輸入錯誤
EXIT_CODE_SAVE_FAILURE = 3  # 儲存失敗
EXIT_CODE_PROCESS_FAILURE = 10  # 過程失敗
EXIT_CODE_UNKNOWN_ERROR = 99  # 未知錯誤


# =========================================================================
# Help 函數
# =========================================================================
def get_workbook():
    """取得【作用中活頁簿】"""
    wb = None
    try:
        wb = xw.apps.active.books.active  # 取得 Excel 作用中的活頁簿檔案
        return wb
    except Exception as e:
        print(f"發生錯誤: {e}")
        print(f"無法找到作用中的 Excel 工作簿: {e}", exc_info=True)
        return EXIT_CODE_NO_FILE


class TestJiKhooDict(unittest.TestCase):

    def test_create_piau_im_dict_from_worksheet(self):
        sheet_name = "缺字表"
        khuat_ji_piau_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, sheet_name)
        han_ji = "圊"
        entries = khuat_ji_piau_dict.get_entry(han_ji)
        print(f"Entries for '{han_ji}': {entries}")

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["tai_gi_im_piau"], "N/A")
        self.assertEqual(entries[0]["hau_ziann_im_piau"], "tsheng1")
        self.assertIn((45, 12), entries[0]["coordinates"])

if __name__ == "__main__":
    # 測試初始化
    wb = get_workbook()
    if not wb:
        print("無法取得作用中的 Excel 工作簿，請確保 Excel 已開啟並有活頁簿檔案。")
        exit(EXIT_CODE_NO_FILE)

    suite = unittest.TestSuite()

    sheet_name = "缺字表"
    suite.addTest(TestJiKhooDict('test_create_piau_im_dict_from_worksheet'))

    runner = unittest.TextTestRunner()
    runner.run(suite)
