import unittest

from a400_反切查拼音 import (
    close_db_connection,
    connect_to_db,
    han_ji_cha_siau_un,
    query_han_ji_piau_im,
)


class TestQueryHanJiPiauIm(unittest.TestCase):
    def setUp(self):
        self.conn, self.cursor = connect_to_db('Kong_Un_V2.db')

    def tearDown(self):
        close_db_connection(self.conn)
        
    def test_query_han_ji_piau_im(self):
        han_ji = '空'
        results = query_han_ji_piau_im(self.cursor, han_ji)
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
            if result.get('小韻字') == '空' and result.get('切語') == '苦紅' and result.get('標音') == 'khong1':
                found = True
                break

        self.assertTrue(found, "未找到預期的讀音資訊")

if __name__ == '__main__':
    unittest.main()
