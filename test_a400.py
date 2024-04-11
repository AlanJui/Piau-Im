import unittest

from a400_反切查拼音 import *


class TestFunctions(unittest.TestCase):
    def setUp(self):
        self.conn, self.cursor = connect_to_db('Kong_Un_V2.db')

    def tearDown(self):
        close_db_connection(self.conn)

    # 以下是函數的測試模版
    # def test_other_function(self):
    #     result = other_function()
    #     self.assertIsNotNone(result)

    def test_query_sip_ngoo_im_siann_bu_tui_ciau_piau(self):
        result = query_sip_ngoo_im_siann_bu_tui_ciau_piau(self.cursor)
        self.assertIsNotNone(result)

    def test_query_sip_ngoo_im_un_bu_tui_ciau_piau(self):
        result = query_sip_ngoo_im_un_bu_tui_ciau_piau(self.cursor)
        self.assertIsNotNone(result)

    def test_query_xiao_yun_cha_xun(self):
        # SQL 查询语句，从"小韻查詢"视图中选择数据
        query = """
        SELECT * 
        FROM 小韻查詢
        LIMIT 1;  -- 只获取一条记录以验证视图是否返回数据
        """
        
        # 使用 cursor 执行查询
        self.cursor.execute(query)
        
        # 获取查询结果
        result = self.cursor.fetchall()
        
        # 验证结果不为空，即视图中有数据
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 0, "小韻查詢视图没有返回数据")

    def test_query_xiao_yun_for_feng(self):
        # 定義 SQL 查詢語句，篩選小韻字為“風”的紀錄
        query = """
        SELECT 切語, 拼音
        FROM 小韻查詢
        WHERE 小韻字 = '風';
        """
        
        # 使用 cursor 執行查詢
        self.cursor.execute(query)
        
        # 獲取查詢結果
        results = self.cursor.fetchall()
        
        # 驗證結果不為空
        self.assertIsNotNone(results)
        self.assertGreater(len(results), 0, "查詢‘風’的小韻沒有返回數據")
        
        # 檢查每一條返回的紀錄
        found = False
        for result in results:
            chiat_gu, phing_im = result
            if chiat_gu == '方戎' and phing_im == 'hiong1':
                found = True
                break
        
        # 驗證是否找到了預期的紀錄
        self.assertTrue(found, "未找到預期的‘切語’為‘方戎’且‘拼音’為‘hiong1’的紀錄")

if __name__ == '__main__':
    unittest.main()
