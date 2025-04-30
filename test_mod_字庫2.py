import unittest

import xlwings as xw

from mod_excel_access import ensure_sheet_exists
from mod_字庫 import JiKhooDict
from mod_標音 import PiauIm, tlpa_tng_han_ji_piau_im


class TestUpdateByPiauImJiKhoo(unittest.TestCase):
    def setUp(self):
        self.wb = xw.Book()
        self.sheet_han_ji = self.wb.sheets.add("漢字注音")
        self.sheet_piau_im = self.wb.sheets.add("標音字庫")

        # 初始化標音字庫內容
        self.sheet_piau_im.range("A1").value = ["漢字", "台語音標", "校正音標", "座標"]
        # self.sheet_piau_im.range("A2").value = ["雨", "hoo7", "ho7", "(5, 4)"]
        self.sheet_piau_im.range("A2").value = ["雨", "hoo7", "ho7", "(5, 4); (5, 4)"]

        # 模擬漢字注音表格中的儲存格：上中下各一列
        self.sheet_han_ji.range((3, 4)).value = ""       # 人工標音（預留）
        self.sheet_han_ji.range((4, 4)).value = ""       # 台語音標（會被覆蓋）
        self.sheet_han_ji.range((5, 4)).value = "雨"     # 漢字
        self.sheet_han_ji.range((6, 4)).value = ""       # 漢字標音（會被覆蓋）

        self.piau_im = PiauIm("河洛話")
        self.piau_im_huat = "TLPA"

    def test_update_by_piau_im(self):
        ji_khoo = JiKhooDict()
        result = ji_khoo.update_by_piau_im_ji_khoo(
            wb=self.wb,
            sheet_name="標音字庫",
            piau_im=self.piau_im,
            piau_im_huat=self.piau_im_huat
        )

        updated_tlpa = self.sheet_han_ji.range((4, 4)).value
        updated_han_ji_piau_im = self.sheet_han_ji.range((6, 4)).value

        self.assertEqual(updated_tlpa, "ho7")
        self.assertIsInstance(updated_han_ji_piau_im, str)
        self.assertTrue(updated_han_ji_piau_im.startswith("h"))
        self.assertEqual(result, 0)

    def tearDown(self):
        self.wb.close()

if __name__ == "__main__":
    unittest.main()
