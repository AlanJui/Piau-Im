import os
import sys
import types
import unittest


FAKE_MODULE_NAMES = [
    "xlwings",
    "mod_excel_access",
    "mod_logging",
    "mod_帶調符音標",
    "mod_標音",
    "mod_程式",
]
ORIGINAL_MODULES = {name: sys.modules.get(name) for name in FAKE_MODULE_NAMES}


def install_fake_module(name, **attrs):
    module = types.ModuleType(name)
    for attr_name, value in attrs.items():
        setattr(module, attr_name, value)
    sys.modules[name] = module
    return module


fake_xlwings = install_fake_module("xlwings")
fake_xlwings.utils = types.SimpleNamespace(col_name=lambda col: chr(64 + col))

install_fake_module("mod_excel_access", get_value_by_name=lambda *args, **kwargs: None)
install_fake_module(
    "mod_logging",
    init_logging=lambda *args, **kwargs: None,
    logging_exc_error=lambda *args, **kwargs: None,
    logging_exception=lambda *args, **kwargs: None,
    logging_process_step=lambda *args, **kwargs: None,
)
install_fake_module(
    "mod_帶調符音標",
    is_han_ji=lambda value: True,
    kam_si_u_tiau_hu=lambda value: False,
    tng_im_piau=lambda value: value,
    tng_tiau_ho=lambda value: value,
)
install_fake_module(
    "mod_標音",
    is_punctuation=lambda value: value in "，。：？！《》·",
    split_tai_gi_im_piau=lambda value: ("", "", ""),
)


class FakeExcelCell:
    def __init__(self, program):
        self.program = program


install_fake_module("mod_程式", ExcelCell=FakeExcelCell, Program=object)

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from a400_製作標音網頁 import CellProcessor


def tearDownModule():
    for name, module in ORIGINAL_MODULES.items():
        if module is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = module


class FakeRange:
    def __init__(self, value):
        self.value = value


class FakeSheet:
    def __init__(self, cells):
        self.cells = cells

    def range(self, coord):
        return FakeRange(self.cells.get(coord))


class TestA400TitleAuthor(unittest.TestCase):
    def test_title_author_reader_includes_last_column(self):
        cells = {
            (5, 4): "《",
            (5, 5): "採",
            (5, 6): "桑",
            (5, 7): "子",
            (5, 8): "·",
            (5, 9): "時",
            (5, 10): "光",
            (5, 11): "只",
            (5, 12): "解",
            (5, 13): "催",
            (5, 14): "人",
            (5, 15): "老",
            (5, 16): "》",
            (5, 17): "北",
            (5, 18): "宋",
            (9, 4): "：",
            (9, 5): "晏",
            (9, 6): "殊",
            (9, 7): "\n",
        }
        sheet = FakeSheet(cells)
        program = types.SimpleNamespace(
            wb=types.SimpleNamespace(sheets={"漢字注音": sheet}),
            line_start_row=3,
            han_ji_row_offset=2,
            start_col=4,
            end_col=18,
            ROWS_PER_LINE=4,
        )
        processor = CellProcessor.__new__(CellProcessor)
        processor.program = program
        processor.generate_ruby_tag = lambda han_ji, tlpa: (
            f"<ruby>{han_ji}</ruby>",
            "",
            "",
        )

        title_html, author_html, next_start_row = (
            processor.generate_title_and_author_with_ruby()
        )

        self.assertIn("採", title_html)
        self.assertIn("北", author_html)
        self.assertIn("宋", author_html)
        self.assertIn("晏", author_html)
        self.assertEqual(next_start_row, 13)


if __name__ == "__main__":
    unittest.main()
