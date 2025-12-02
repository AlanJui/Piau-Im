"""
測試 mod_ca_ji_tian 與 mod_標音 的整合
"""

from mod_ca_ji_tian import HanJiTian
from mod_標音 import PiauIm, ca_ji_kiat_ko_tng_piau_im

# 初始化
ji_tian = HanJiTian("Ho_Lok_Ue.db")
piau_im = PiauIm(han_ji_khoo="河洛話")

print("=" * 70)
print("測試整合功能：查詢漢字 → 轉換音標")
print("=" * 70)

# 測試查詢和轉換
test_chars = ["東", "西", "南"]
for han_ji in test_chars:
    print(f"\n查詢漢字：{han_ji}")

    # 查詢白話音
    result = ji_tian.han_ji_ca_piau_im(han_ji, ue_im_lui_piat="白話音")

    if result:
        print(f"  查詢結果：{result[0]}")

        # 轉換為標音
        tai_gi_im_piau, han_ji_piau_im = ca_ji_kiat_ko_tng_piau_im(
            result=result,
            han_ji_khoo="河洛話",
            piau_im=piau_im,
            piau_im_huat="方音符號"
        )

        print(f"  台語音標：{tai_gi_im_piau}")
        print(f"  方音符號：{han_ji_piau_im}")
    else:
        print(f"  查無資料")

print("\n" + "=" * 70)
print("測試完成")
print("=" * 70)
