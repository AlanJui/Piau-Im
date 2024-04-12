import unittest

from mod_廣韻 import (
    cha_ciat_gu_ha_ji,
    cha_ciat_gu_siong_ji,
    cha_siann_bu_tui_ciau_piau,
    cha_un_bu_tui_ciau_piau,
    close_db_connection,
    connect_to_db,
    han_ji_cha_piau_im,
    han_ji_cha_siau_un,
    query_table_by_field,
    query_table_by_id,
)


class TestQueryHanJiPiauIm(unittest.TestCase):
    def setUp(self):
        self.conn, self.cursor = connect_to_db('Kong_Un_V2.db')

    def tearDown(self):
        close_db_connection(self.conn)
        
    def test_cha_siann_bu_tui_ciau_piau(self):
        result = cha_siann_bu_tui_ciau_piau(self.cursor)
        self.assertIsNotNone(result)


    def test_query_cha_siau_un(self):
        """測試能否操作資料庫之檢視(View)
        """

        # SQL 查询语句，从"小韻查詢"视图中选择数据
        query = """
        SELECT * 
        FROM 小韻檢視
        LIMIT 1;  -- 只获取一条记录以验证视图是否返回数据
        """
        
        # 使用 cursor 执行查询
        self.cursor.execute(query)
        
        # 获取查询结果
        result = self.cursor.fetchall()
        
        # 验证结果不为空，即视图中有数据
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 0, "小韻查詢视图没有返回数据")


    def test_cha_un_bu_tui_ciau_piau(self):
        result = cha_un_bu_tui_ciau_piau(self.cursor)
        self.assertIsNotNone(result)


    def test_query_table_by_id(self):
        table_name = '切語上字表'
        fields = ['識別號', '發音部位', '聲母', '清濁', '發送收', '聲母拼音碼', '國際音標', '切語上字', '備註']
        id = 1
        result = query_table_by_id(self.cursor, table_name, fields, id)
        self.assertIsNotNone(result)
        

    def test_query_table_by_field(self):
        table_name = '小韻表'
        fields = [
            '識別號', '上字表識別號', '下字表識別號', '切語', '拼音', '小韻字',
            '目次編碼', '小韻字序號', '小韻字集', '字數',
            '聲母', '聲母拼音碼', '發音部位', '清濁', '發送收',
            '韻母', '韻母拼音碼', '調', '調號',
            '備註', '原有備註', '異體字', '其它備註',
        ]
        query_field = '小韻字集'
        siau_un_ji = '中'
        result = query_table_by_field(self.cursor, table_name, fields, query_field, siau_un_ji)
        self.assertIsNotNone(result)

        # 檢查結果中是否包含 "中"
        found = False
        for row in result:
            if '中' in row.get('小韻字集', ''):
                found = True
                break

        self.assertTrue(found, "在 '小韻字集' 欄位中未找到 '中'")


    def test_han_ji_cha_siau_un(self):
        # 定義 SQL 查詢語句，篩選小韻字為“風”的紀錄
        query = """
        SELECT 小韻切語, 小韻標音
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


    def test_han_ji_cha_piau_im(self):
        han_ji = '空'
        results = han_ji_cha_piau_im(self.cursor, han_ji)
        self.assertGreater(len(results), 0, "未查詢到‘空’的讀音資訊")
        # 驗證是否包含特定的讀音資訊
        expected = {'小韻字': '空', '切語': '苦紅', '標音': 'khong1'}
        self.assertIn(expected, results, "未找到預期的讀音資訊")


    def test_han_ji_cha_siau_un(self):
        han_ji = '空'
        results = han_ji_cha_siau_un(self.cursor, han_ji)
        self.assertGreater(len(results), 0, "未查詢到‘空’的讀音資訊")

        # 检查是否存在符合条件的记录
        found = False
        for result in results:
            if result.get('小韻字') == '空' and result.get('小韻切語') == '苦紅' and result.get('小韻標音') == 'khong1':
                found = True
                break

        self.assertTrue(found, "未找到預期的讀音資訊")


    def test_han_ji_cha_piau_im(self):
        han_ji = '空'
        results = han_ji_cha_piau_im(self.cursor, han_ji)
        self.assertGreater(len(results), 0, "未查詢到‘空’的讀音資訊")

        # 检查是否存在符合条件的记录
        found = False
        for result in results:
            if result.get('小韻字') == '空' and result.get('切語') == '苦紅' and result.get('漢字標音') == 'khong1':
                found = True
                break

        self.assertTrue(found, "未找到預期的讀音資訊")
        

    def test_cha_ciat_gu_siong_ji(self):
        siong_ji = '魚'
        results = cha_ciat_gu_siong_ji(self.cursor, siong_ji)
        self.assertIsNotNone(results)

        # 检查是否存在符合条件的记录
        found = False
        for result in results:
            if result.get('聲母') == '疑' and result.get('聲母拼音碼') == 'g':
                found = True
                break

        self.assertTrue(found, "未找到預期的讀音資訊")


    def test_cha_ciat_gu_ha_ji(self):
        ha_ji = '宗'
        results = cha_ciat_gu_ha_ji(self.cursor, ha_ji)
        self.assertIsNotNone(results)

        # 检查是否存在符合条件的记录
        found = False
        for result in results:
            if result.get('韻母') == '冬開1' and result.get('韻母拼音碼') == 'ong':
                found = True
                break

        self.assertTrue(found, "未找到預期的讀音資訊")


if __name__ == '__main__':
    unittest.main()
