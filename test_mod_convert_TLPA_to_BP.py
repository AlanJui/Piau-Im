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
    # 台羅音標（TL）
    ("熱", "jiat8", "zziat8" ), # ㆢ：ji → zz+i
    ("熱", "juah8", "zzuah8"),  # ㆡ：j -> zz
    ("曾", "tsan1", "zan1"),    # ㄗ：ts -> z
    ("尖", "tsiam1", "ziam1"),  # ㄐ：ts+i -> z+i
    ("出", "tshut4", "cut7"),   # ㄘ：tsh -> c
    ("手", "tshiu2", "ciu3"),   # ㄑ：tsh+i -> c+i
    #----------------------------------------------
    # 台語音標（TLPA）
    ("邊", "pian1", "bian1"),   # ㄅ：p -> b
    ("文", "bun5", "bbun2"),    # ㆠ：b -> bb
    ("頗", "pho7", "po6"),      # ㄆ：ph -> p
    ("毛", "moo7", "bbnoo6"),   # ㄇ：m -> bbn
    #----------------------------------------------
    ("地", "te2", "de3"),       # ㄉ：t -> d
    ("他", "thann1", "tna1"),   # ㄊ：th -> t
    ("耐", "nai2", "lnai3"),    # ㄋ：n -> ln
    ("柳", "liu2", "liu3"),     # ㄌ：l -> l
    #----------------------------------------------
    ("曾", "zan1", "zan1"),     # ㄗ：z -> z
    ("熱", "juah8", "zzuah8"),  # ㆡ：j -> zz
    ("出", "cut4", "cut7"),     # ㄘ：c -> c
    ("衫", "sann1", "sna1"),    # ㄙ：s -> s
    #----------------------------------------------
    ("尖", "ziam1", "ziam1"),   # ㄐ：z+i -> z+i
    ("入", "jip4", "zzip7" ),   # ㆢ：j+i → zz+i
    ("手", "ciu2", "ciu3"),     # ㄑ：c+i -> c+i
    ("寫", "sia2", "sia3"),     # ㄒ：s+i -> s+i
    #----------------------------------------------
    ("求", "kiu5", "giu2"),     # ㄍ：k -> g
    ("語", "gi2", "ggi3"),      # ㆣ：g -> gg
    ("去", "khi2", "ki3"),      # ㄎ：kh -> k
    ("雅", "nga2", "ggna3"),    # ㄫ：ng -> ggn
    #----------------------------------------------
    # 閩拼專用測試案例：零聲母 + i/u 規則
    ("依", "i1", "yi1"),
    ("因", "in1", "yin1"),
    ("鴉", "ia1", "ya1"),
    ("煙", "ian1", "yan1"),
    ("用", "iong7", "yong6"),
    ("烏", "u1", "wu1"),
    ("運", "un7", "wun6"),
    ("媧", "ua1", "wa1"),
    ("彎", "uan1", "wan1"),
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
