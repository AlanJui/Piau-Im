# 撰寫 test_mod_字庫.py 的測試案例，用於驗證 remove_coordinate() 方法行為
import unittest
from copy import deepcopy
from types import SimpleNamespace

# 假設這是從 mod_字庫.py 匯入的類別
from mod_字庫 import JiKhooDict


class TestJiKhooDict(unittest.TestCase):

    def setUp(self):
        self.dict_obj = JiKhooDict()
        self.han_ji = "行"
        self.piau = "hang5"
        self.coord1 = (9, 7)
        self.coord2 = (10, 8)
        self.dict_obj.add_entry(self.han_ji, self.piau, "N/A", self.coord1)
        self.dict_obj.add_entry(self.han_ji, self.piau, "N/A", self.coord2)

    def test_remove_existing_coordinate(self):
        self.dict_obj.remove_coordinate(self.han_ji, self.piau, self.coord1)
        entry = self.dict_obj.get_entry(self.han_ji)[0]
        self.assertNotIn(self.coord1, entry["coordinates"])
        self.assertIn(self.coord2, entry["coordinates"])

    def test_remove_all_coordinates_results_in_removal(self):
        self.dict_obj.remove_coordinate(self.han_ji, self.piau, self.coord1)
        self.dict_obj.remove_coordinate(self.han_ji, self.piau, self.coord2)
        self.assertEqual(self.dict_obj.ji_khoo_dict[self.han_ji], [])

    def test_remove_nonexistent_coordinate_does_nothing(self):
        before = deepcopy(self.dict_obj.ji_khoo_dict)
        self.dict_obj.remove_coordinate(self.han_ji, self.piau, (99, 99))
        self.assertEqual(before, self.dict_obj.ji_khoo_dict)

    def test_remove_coordinate_invalid_han_ji(self):
        # 不會 raise error，只是 silently pass
        try:
            self.dict_obj.remove_coordinate("不存在字", "xxx1", (1, 1))
        except Exception as e:
            self.fail(f"unexpected exception: {e}")

unittest.TextTestRunner().run(unittest.TestLoader().loadTestsFromTestCase(TestJiKhooDict))
