"""
測試自動製作打字練習表的功能
"""

from auto_typing_practice import decompose_pronunciation, get_tone_key_mapping


def test_decompose_pronunciation():
    """
    測試拼音分解功能
    """
    print("=== 測試拼音分解功能 ===")

    # 測試案例
    test_cases = [
        # 羅馬拼音測試
        ("tong1", ["t", "o", "n", "g", ";"]),
        ("tong2", ["t", "o", "n", "g", "\\"]),
        ("tong3", ["t", "o", "n", "g", "_"]),
        ("tong4", ["t", "o", "k", "["]),
        ("tong5", ["t", "o", "n", "g", "/"]),
        ("tong7", ["t", "o", "n", "g", "-"]),
        ("tong8", ["t", "o", "k", "]"]),
        ("su5", ["s", "u", "/"]),

        # 注音符號測試
        ("ㄌㄧㄤˊ", ["ㄌ", "ㄧ", "ㄤ", "6"]),
        ("ㄉㄧㆻ˙", ["ㄉ", "ㄧ", "ㆻ", " "]),
        ("ㄙㄨ", ["ㄙ", "ㄨ", " "]),  # 沒有聲調符號，預設陰平
    ]

    for i, (input_str, expected) in enumerate(test_cases, 1):
        result = decompose_pronunciation(input_str)
        status = "✓" if result == expected else "✗"
        print(f"測試 {i:2d}: {input_str:8s} → {result} {status}")
        if result != expected:
            print(f"         期望: {expected}")
            print(f"         實際: {result}")
        print()


def test_tone_mapping():
    """
    測試聲調對照表
    """
    print("=== 測試聲調對照表 ===")

    roman_map, bopomofo_map = get_tone_key_mapping()

    print("羅馬拼音聲調對照:")
    for tone, key in roman_map.items():
        print(f"  {tone} → {repr(key)}")

    print("\n注音符號聲調對照:")
    for tone, key in bopomofo_map.items():
        print(f"  {repr(tone)} → {repr(key)}")


if __name__ == "__main__":
    test_tone_mapping()
    print()
    test_decompose_pronunciation()