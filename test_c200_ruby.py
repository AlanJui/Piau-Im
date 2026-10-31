"""c200：漢字標音儲存格寫入 <rt>，方音符號寫入 <rtc>。"""

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


def _is_han_ji(value):
    if not isinstance(value, str) or len(value) != 1:
        return False
    return "\u4e00" <= value <= "\u9fff"


def _is_punctuation(value):
    return isinstance(value, str) and len(value) == 1 and value in "，。：？！《》·"


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
    is_han_ji=_is_han_ji,
    kam_si_u_tiau_hu=lambda value: False,
    tng_im_piau=lambda value: value,
    tng_tiau_ho=lambda value: value,
)
install_fake_module(
    "mod_標音",
    is_punctuation=_is_punctuation,
    split_tai_gi_im_piau=lambda value: ("t", "uan", "3"),
)


class FakeExcelCell:
    def __init__(self, program):
        self.program = program


install_fake_module("mod_程式", ExcelCell=FakeExcelCell, Program=object)

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from c200_使用漢字標音儲存格之變調標注製成網頁 import (  # noqa: E402
    CellProcessor,
    _build_head_extra,
    _clean_named_value,
)


def tearDownModule():
    for name, module in ORIGINAL_MODULES.items():
        if module is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = module


class FakeRange:
    def __init__(self, value, color=None):
        self.value = value
        self.color = color


class FakeSheet:
    def __init__(self, cells):
        self.cells = cells

    def range(self, coord):
        return FakeRange(*self.cells.get(coord, (None, None)))


class TestC200Ruby(unittest.TestCase):
    def _processor(self):
        processor = CellProcessor.__new__(CellProcessor)
        processor.program = types.SimpleNamespace(
            piau_im=types.SimpleNamespace(
                han_ji_piau_im_tng_huan=lambda *args: "ㄉㄨㄢ˪"
            )
        )
        return processor

    def test_punctuation_is_span_without_rt(self):
        tag, siong, zian = self._processor().generate_ruby_tag("。", "", "。")
        self.assertIn("<span>。</span>", tag)
        self.assertNotIn("<rt>", tag)
        self.assertNotIn("<rtc>", tag)
        self.assertEqual(siong, "")
        self.assertEqual(zian, "")

    def test_han_ji_piau_im_goes_to_rt_and_tps_to_rtc(self):
        tag, siong, zian = self._processor().generate_ruby_tag(
            "斷",
            "tuan3",
            han_ji_piau_im="duan⁵⁻³",
            pian_tiau=True,
        )
        self.assertIn('data-tlpa="tuan3"', tag)
        self.assertIn('data-pian-tiau="true"', tag)
        self.assertIn("<rt>duan⁵⁻³</rt>", tag)
        self.assertIn("<rtc>ㄉㄨㄢ˪</rtc>", tag)
        self.assertEqual(siong, "duan⁵⁻³")
        self.assertEqual(zian, "ㄉㄨㄢ˪")

    def test_punctuation_in_piau_im_cell_is_not_copied_to_rt(self):
        tag, siong, zian = self._processor().generate_ruby_tag(
            "斷",
            "tuan3",
            han_ji_piau_im="，",
            pian_tiau=False,
        )
        self.assertNotIn("<rt>", tag)
        self.assertIn("<rtc>ㄉㄨㄢ˪</rtc>", tag)
        self.assertEqual(siong, "")
        self.assertEqual(zian, "ㄉㄨㄢ˪")
        self.assertIn('data-pian-tiau="false"', tag)

    def test_title_reader_uses_han_ji_piau_im_cell(self):
        gray = (232, 232, 232)
        cells = {
            (5, 4): ("《", None),
            (5, 5): ("斷", gray),
            (4, 5): ("tuan3", None),
            (6, 5): ("duan⁵⁻³", None),
            (5, 6): ("》", None),
            (5, 7): ("：", None),
            (5, 8): ("虹", None),
            (4, 8): ("hong5", None),
            (6, 8): ("hong²", None),
            (5, 9): ("\n", None),
        }
        sheet = FakeSheet(cells)
        program = types.SimpleNamespace(
            wb=types.SimpleNamespace(sheets={"漢字注音": sheet}),
            line_start_row=3,
            han_ji_row_offset=2,
            start_col=4,
            end_col=9,
            ROWS_PER_LINE=4,
        )
        processor = self._processor()
        processor.program.wb = program.wb
        processor.program.line_start_row = program.line_start_row
        processor.program.han_ji_row_offset = program.han_ji_row_offset
        processor.program.start_col = program.start_col
        processor.program.end_col = program.end_col
        processor.program.ROWS_PER_LINE = program.ROWS_PER_LINE

        title_html, author_html, next_start_row = (
            processor.generate_title_and_author_with_ruby()
        )

        self.assertIn('class="title_mark"', title_html)
        self.assertIn("<rt>duan⁵⁻³</rt>", title_html)
        self.assertIn('data-pian-tiau="true"', title_html)
        self.assertIn("<rt>hong²</rt>", author_html)
        self.assertIn("虹", author_html)
        self.assertNotIn("<rt>：</rt>", title_html + author_html)
        self.assertEqual(next_start_row, 9)

    def test_head_extra_does_not_enable_pian_tiau_rewrite(self):
        self.assertEqual(_clean_named_value(None), "")
        self.assertEqual(_clean_named_value("None"), "")

        class FakeBook:
            def __init__(self):
                self.names = {}

        # _build_head_extra 只透過 get_value_by_name 讀 TITLE／IMAGE_URL，
        # 測試用的假模組固定回傳 None。
        head = _build_head_extra(FakeBook(), "閩拼調符")
        self.assertIn('name="使用變調標音格式" content=""', head)
        self.assertIn('name="右邊標音" content="方音符號"', head)
        self.assertIn('name="上邊標音" content="閩拼調符"', head)
        self.assertNotIn('content="是"', head)


if __name__ == "__main__":
    unittest.main()
