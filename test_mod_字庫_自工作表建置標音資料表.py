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

    def test_add_entry(self):
        ji_khoo = JiKhooDict()
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 14))
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 17))
        ji_khoo.add_entry("雨", "u2", "N/A", (5, 18))

        entries = ji_khoo.get_entry("雨")
        self.assertEqual(len(entries), 2)
        self.assertEqual(entries[0]["tai_gi_im_piau"], "hoo7")
        self.assertIn((5, 14), entries[0]["coordinates"])
        self.assertIn((5, 18), entries[1]["coordinates"])

    def test_multiple_readings(self):
        ji_khoo = JiKhooDict()
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 14))
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 17))
        ji_khoo.add_entry("雨", "u2", "N/A", (57, 18))

        entries = ji_khoo.get_entry("雨")
        self.assertEqual(len(entries), 2)
        self.assertEqual(entries[0]["tai_gi_im_piau"], "hoo7")
        self.assertEqual(entries[1]["tai_gi_im_piau"], "u2")

    def test_write_to_excel_sheet(self):
        wb = xw.Book()
        sheet_name = "標音字庫測試用"

        # 建立測試用【標音資料表】，並寫入工作表
        ji_khoo = JiKhooDict()
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 14))
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 17))
        ji_khoo.add_entry("雨", "u2", "N/A", (57, 18))

        ji_khoo.write_to_excel_sheet(wb, sheet_name)

        # 自工作表讀取資料，驗證寫入的資料是否正確
        sheet = wb.sheets[sheet_name]
        data = sheet.range("A2").expand("table").value

        # 驗證寫入的資料是否正確
        self.assertEqual(len(data), 2)

        self.assertEqual(data[0][0], "雨")
        self.assertEqual(data[0][1], "hoo7")

        self.assertEqual(data[1][0], "雨")
        self.assertEqual(data[1][1], "u2")

        wb.close()

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

    def test_list_all_items_in_dict(self):
        # 測試初始化
        wb = get_workbook()
        if not wb:
            print("無法取得作用中的 Excel 工作簿，請確保 Excel 已開啟並有活頁簿檔案。")
            exit(EXIT_CODE_NO_FILE)

        sheet_name = "缺字表"
        han_ji_piau_im_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, sheet_name)
        # 讀取【標音字庫/人工標音字庫】工作表所製成之【漢字標音資料表】的【總筆數】
        total_entries = len(han_ji_piau_im_dict)
        print(f"從【{sheet_name}】工作表讀取到 {total_entries} 筆資料。")

        # 自【漢字標音資料表】中遍歷【標音工作表】的每一筆資料（row）
        for idx, item in enumerate(han_ji_piau_im_dict, start=1):
            han_ji = item.get("漢字")
            tai_gi_im_piau = item.get("台語音標")
            hau_ziann_im_piau = item.get("校正音標")
            zo_piau = item.get("座標")
            print(f"{idx}. 【{han_ji}】：台語音標=【{tai_gi_im_piau}】、校正音標=【{hau_ziann_im_piau}】、座標={zo_piau}")

    def test_zit_ji_to_im(self):
        wb = xw.Book()
        sheet_name = "標音字庫測試用"

        # 建立測試用【標音資料表】，並寫入工作表
        ji_khoo = JiKhooDict()
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 14))
        ji_khoo.add_entry("雨", "hoo7", "N/A", (5, 17))
        ji_khoo.add_entry("雨", "u2", "N/A", (57, 18))

        ji_khoo.write_to_excel_sheet(wb, sheet_name)

        # 自工作表讀取資料，驗證寫入的資料是否正確
        try:
            # 自【標音字庫/人工標音字庫】工作表，産製 han_ji_piau_im_dict
            han_ji_piau_im_dict = []
            han_ji_piau_im_dict = JiKhooDict.create_ji_khoo_dict_from_sheet(wb, sheet_name)
        except Exception as e:
            print(f"無法從《{sheet_name}》工作表讀取資料，產生【漢字標音資料表】（han_ji_piau_im_dict）。")
            print(f"問題詳述：{e}")
            return EXIT_CODE_INVALID_INPUT

        # 讀取【標音字庫/人工標音字庫】工作表所製成之【漢字標音資料表】的【總筆數】
        total_entries = len(han_ji_piau_im_dict)
        print(f"從【{sheet_name}】工作表讀取到 {total_entries} 筆資料。")

        # 自【漢字標音資料表】中遍歷【標音工作表】的每一筆資料（row）
        for idx, item in enumerate(han_ji_piau_im_dict, start=1):
            han_ji = item.get("漢字")
            tai_gi_im_piau = item.get("台語音標")
            hau_ziann_im_piau = item.get("校正音標")
            zo_piau = item.get("座標")
            print(f"{idx}. 【{han_ji}】：台語音標=【{tai_gi_im_piau}】、校正音標=【{hau_ziann_im_piau}】、座標={zo_piau}")

        # 驗證寫入的資料是否正確
        self.assertEqual(len(han_ji_piau_im_dict), 2)

        # 使用索引方式訪問 han_ji_piau_im_dict 中的項目，驗證資料是否正確
        self.assertEqual(han_ji_piau_im_dict[0]["漢字"], "雨")
        self.assertEqual(han_ji_piau_im_dict[0]["台語音標"], "hoo7")

        self.assertEqual(han_ji_piau_im_dict[1]["漢字"], "雨")
        self.assertEqual(han_ji_piau_im_dict[1]["台語音標"], "u2")

        # 使用 get_entry 方法訪問 han_ji_piau_im_dict 中的項目，驗證資料是否正確
        self.assertEqual(han_ji_piau_im_dict["雨"][0]["tai_gi_im_piau"], "hoo7")
        self.assertEqual(han_ji_piau_im_dict["雨"][1]["tai_gi_im_piau"], "u2")

        wb.close()


if __name__ == "__main__":
    suite = unittest.TestSuite()

    #========================================================================
    # 基本功能測試
    #========================================================================
    # suite.addTest(TestJiKhooDict('test_add_entry'))
    # suite.addTest(TestJiKhooDict('test_multiple_readings'))
    # suite.addTest(TestJiKhooDict('test_write_to_excel_sheet'))

    #========================================================================
    # 進階功能測試
    #========================================================================
    sheet_name = "缺字表"
    # suite.addTest(TestJiKhooDict('test_create_piau_im_dict_from_worksheet'))
    # 一字多音測試
    suite.addTest(TestJiKhooDict('test_zit_ji_to_im'))

    runner = unittest.TextTestRunner()
    runner.run(suite)
