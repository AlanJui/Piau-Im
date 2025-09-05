# -*- coding: utf-8 -*-
"""
單元測試：TLPA → BP 轉換（零聲母 + i/u 規則驗證）

執行方式：
    python test_convert_TLPA_to_MPS2.py
"""
import importlib.util
import sys
import unittest
from pathlib import Path

# MODULE_PATH = Path("/mnt/data/mod_convert_TLPA_to_MPS2.py")
# MODULE_PATH = Path("./mod_convert_TLPA_to_MPS2.py")
MODULE_PATH = Path("c:/work/Piau-Im/mod_convert_TLPA_to_MPS2.py")

# 將模組所在資料夾加入 sys.path 以便 import
if str(MODULE_PATH.parent) not in sys.path:
    sys.path.insert(0, str(MODULE_PATH.parent))

# 動態載入模組
spec = importlib.util.spec_from_file_location("mod_convert_TLPA_to_MPS2", str(MODULE_PATH))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)  # type: ignore

# 嘗試取得轉換函式
convert_func = None
for name in ["convert_TLPA_to_MPS2", "convert_tlpa_to_bp", "tlpa_to_bp"]:
    if hasattr(mod, name):
        convert_func = getattr(mod, name)
        break

if not callable(convert_func):
    raise AttributeError("在 mod_convert_TLPA_to_MPS2.py 找不到可呼叫的函式：convert_TLPA_to_MPS2（或可能名稱不同）")

TEST_CASES = [
    # 台羅音標（TL）
    ("熱", "jiat8", "jjiat8" ), # ㆢ：ji → jj+i
    ("入", "jip4", "jjip4" ),   # ㆢ：ji → jj+i
    ("熱", "juah8", "zzuah8"),  # ㆡ：j -> zz
    ("曾", "tsan1", "zan1"),    # ㄗ：ts -> z
    ("尖", "tsiam1", "ziam1"),  # ㄐ：ts+i -> z+i
    ("出", "tshut4", "cut4"),   # ㄘ：tsh -> c
    ("手", "tshiu2", "ciu2"),   # ㄑ：tsh+i -> c+i
    ("衫", "sann1", "sann1"),   # ㄙ：s -> s
    ("寫", "sia2", "sia2"),     # ㄒ：s+i -> s+i
    # 台語音標（TL）
    ("曾", "zan1", "zan1"),     # ㄗ：z -> z
    ("尖", "ziam1", "ziam1"),   # ㄐ：z+i -> z+i
    ("出", "cut4", "cut4"),     # ㄘ：c -> c
    ("手", "ciu2", "ciu2"),     # ㄑ：c+i -> c+i
    # 閩拼專用測試案例：零聲母 + i/u 規則
    # ("依", "i1", "yi1"),
    # ("因", "in1", "yin1"),
    # ("鴉", "ia1", "ya1"),
    # ("煙", "ian1", "yan1"),
    # ("用", "iong7", "yong7"),
    # ("烏", "u1", "wu1"),
    # ("運", "un7", "wun7"),
    # ("媧", "ua1", "wa1"),
    # ("彎", "uan1", "wan1"),
]

class TestConvertTLPAtoBP(unittest.TestCase):
    def test_cases(self):
        print("\n開始執行 TLPA → BP 轉換測試...")
        for han_ji, tlpa, expected in TEST_CASES:
            with self.subTest(han_ji=han_ji, tlpa=tlpa):
                got = convert_func(tlpa)
                self.assertEqual(expected, got, f"{han_ji}：{tlpa} → {got}（應為 {expected}）")
                print(f"測試通過：{han_ji}：{tlpa} → {got}")

if __name__ == "__main__":
    unittest.main(verbosity=2)
