
import os
import sys
import unittest
from unittest.mock import MagicMock, Mock, patch

# Mock xlwings before importing code if it's not installed
if 'xlwings' not in sys.modules:
    sys.modules['xlwings'] = MagicMock()

# Mock openpyxl since only part of mod_file_access is used or we want to avoid import errors
if 'openpyxl' not in sys.modules:
    sys.modules['openpyxl'] = MagicMock()

# Ensure the module can be imported
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from a210_漢字以人工標音處理作業 import CellProcessor


class TestCellProcessorRepro(unittest.TestCase):
    def setUp(self):
        # Patch dependencies
        self.init_ji_khoo_patcher = patch('mod_程式.ExcelCell._initialize_ji_khoo', return_value={})
        self.mock_init_ji_khoo = self.init_ji_khoo_patcher.start()

        self.is_han_ji_patcher = patch('a210_漢字以人工標音處理作業.is_han_ji', return_value=True)
        self.mock_is_han_ji = self.is_han_ji_patcher.start()

        # Mock Program
        self.mock_program = Mock()
        self.mock_program.db_name = ":memory:"
        self.mock_program.piau_im = Mock()
        self.mock_program.piau_im_huat = "閩拼"

        # Mock the dictionary with the actual structure used in JiKhooDict (List of Dicts)
        self.mock_program.jin_kang_piau_im_ji_khoo_dict = {
            # "而": [{"tai_gi_im_piau": "ji5", "kenn_ziann_im_piau": "N/A", "coordinates": ["(1,1)"]}]
            "於": [{"tai_gi_im_piau": "u5", "hau_ziann_im_piau": "N/A", "coordinates": ["(1,1)"]}]
        }

        self.cell_processor = CellProcessor(program=self.mock_program)

    def tearDown(self):
        self.init_ji_khoo_patcher.stop()
        self.is_han_ji_patcher.stop()

    @patch('a210_漢字以人工標音處理作業.split_tai_gi_im_piau')
    def test_manual_annotation_equal_sign(self, mock_split):
        """
        測試當人工標音為 '=' 時，是否能正確從字庫結構（List of Dicts）中取得音標字串。
        """
        # Arrange
        # mock_split.return_value = ("j", "i", "5")
        mock_split.return_value = ("", "u", "5")

        mock_cell = MagicMock()
        mock_cell.value = "於"  # 漢字

        def offset_side_effect(row_offset, col_offset):
            m = MagicMock()
            if row_offset == -2:
                m.value = "="
            return m
        mock_cell.offset.side_effect = offset_side_effect

        # Act
        self.cell_processor._process_cell(cell=mock_cell, row=21, col=4)

        # Assert
        # 驗證傳遞給 split_tai_gi_im_piau 的參數是否為純字串 "ji5"
        # 修正前：會傳入 "{'tai_gi_im_piau': 'ji5', ...}" 字串
        mock_split.assert_called()
        arg_passed = mock_split.call_args[0][0]

        # self.assertEqual(arg_passed, "ji5", f"預期取得 'ji5'，但得到: {arg_passed}")
        self.assertEqual(arg_passed, "u5", f"預期取得 'u5'，但得到: {arg_passed}")

if __name__ == "__main__":
    unittest.main()
