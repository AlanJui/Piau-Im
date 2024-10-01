# 測試執行方法： python -m unittest test_mod_雅俗通.py
import sqlite3
import unittest

from mod_雅俗通 import han_ji_cha_piau_im, split_cu_im  # 假設你將函數保存在 'your_module.py' 中


class TestHanJiChaPiauIm(unittest.TestCase):
    
    @classmethod
    def setUpClass(cls):
        # 在所有測試開始前，連接資料庫
        cls.conn = sqlite3.connect('Nga_Siok_Thong_Sip_Ngoo_Im.db')  # 替換為實際資料庫路徑
        cls.cursor = cls.conn.cursor()
    
    @classmethod
    def tearDownClass(cls):
        # 在所有測試結束後，關閉資料庫連接
        cls.conn.close()
    
    def test_han_ji_cha_piau_im(self):
        # 測試漢字查詢功能
        han_ji = '不'
        result = han_ji_cha_piau_im(self.cursor, han_ji)
        self.assertEqual(result[0]['十五音聲母'], '邊', "轉換錯誤！")
        self.assertEqual(result[0]['十五音韻母'], '君', "轉換錯誤！")
        self.assertEqual(result[0]['十五音聲調'], '上入', "轉換錯誤！")
        self.assertEqual(result[0]['八聲調'], 4, "轉換錯誤！")
        self.assertEqual(result[0]['聲母台語音標'], 'p', "轉換錯誤！")
        self.assertEqual(result[0]['韻母台語音標'], 'ut', "轉換錯誤！")
        self.assertEqual(result[0]['聲母方音符號'], 'ㄅ', "轉換錯誤！")
        self.assertEqual(result[0]['韻母方音符號'], 'ㄨㆵ', "轉換錯誤！")

    def test_split_cu_im(self):
        # 測試 `split_cu_im` 函數
        cu_im = "put4"
        result = split_cu_im(cu_im)
        self.assertEqual(result[0], 'p', "聲母錯誤！")
        self.assertEqual(result[1], 'ut', "韻母錯誤！")
        self.assertEqual(result[2], '4', "聲調錯誤！")


if __name__ == '__main__':
    unittest.main()
