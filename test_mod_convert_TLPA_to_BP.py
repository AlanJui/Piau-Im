# -*- coding: utf-8 -*-
"""
單元測試：TLPA → BP 轉換（零聲母 + i/u 規則驗證）

執行方式：
    python test_convert_TLPA_to_BP.py
"""
import importlib.util
import sys
import unittest
from pathlib import Path

# MODULE_PATH = Path("/mnt/data/mod_convert_TLPA_to_BP.py")
# MODULE_PATH = Path("./mod_convert_TLPA_to_BP.py")
MODULE_PATH = Path("c:/work/Piau-Im/mod_convert_TLPA_to_BP.py")

# 將模組所在資料夾加入 sys.path 以便 import
if str(MODULE_PATH.parent) not in sys.path:
    sys.path.insert(0, str(MODULE_PATH.parent))

# 動態載入模組
spec = importlib.util.spec_from_file_location("mod_convert_TLPA_to_BP", str(MODULE_PATH))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)  # type: ignore

# 嘗試取得轉換函式
convert_func = None
for name in ["convert_TLPA_to_BP", "convert_tlpa_to_bp", "tlpa_to_bp"]:
    if hasattr(mod, name):
        convert_func = getattr(mod, name)
        break

if not callable(convert_func):
    raise AttributeError("在 mod_convert_TLPA_to_BP.py 找不到可呼叫的函式：convert_TLPA_to_BP（或可能名稱不同）")

TEST_CASES = [
    ("依", "i1", "yi1"),
    ("因", "in1", "yin1"),
    ("鴉", "ia1", "ya1"),
    ("煙", "ian1", "yan1"),
    ("用", "iong7", "yong7"),
    ("烏", "u1", "wu1"),
    ("運", "un7", "wun7"),
    ("媧", "ua1", "wa1"),
    ("彎", "uan1", "wan1"),
]

class TestConvertTLPAtoBP(unittest.TestCase):
    def test_cases(self):
        for han_ji, tlpa, expected in TEST_CASES:
            with self.subTest(han_ji=han_ji, tlpa=tlpa):
                got = convert_func(tlpa)
                self.assertEqual(expected, got, f"{han_ji}：{tlpa} → {got}（應為 {expected}）")
                print(f"測試通過：{han_ji}：{tlpa} → {got}")

if __name__ == "__main__":
    unittest.main(verbosity=2)
