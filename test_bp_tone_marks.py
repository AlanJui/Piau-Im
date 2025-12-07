"""測試 convert_TLPA_to_BP_with_tone_marks 函數"""

from mod_BP_tng_huan_ping_im import convert_TLPA_to_BP_with_tone_marks

# 測試案例
test_cases = [
    ("kun3", "滾"),
    ("iao2", "腰"),
    ("uai3", "歪"),
    ("i1", "伊"),
    ("im1", "音"),
    ("u2", "有"),
    ("un1", "溫"),
    ("a2", "阿"),
    ("iong5", "央"),
    ("uan1", "彎"),
]

print("=" * 70)
print("測試 convert_TLPA_to_BP_with_tone_marks")
print("=" * 70)

for tlpa, han_ji in test_cases:
    result = convert_TLPA_to_BP_with_tone_marks(tlpa)
    print(f"{han_ji:4s} {tlpa:8s} → {result}")

print("=" * 70)
