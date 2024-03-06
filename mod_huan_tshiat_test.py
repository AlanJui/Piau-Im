# mod_huan_tshiat_test.py 內容

from mod_huan_tshiat import siong_ji_tsa_siann_bu, ha_ji_tsa_un_bu

# 測試 siong_ji_tsa_siann_bu 函數
siong_ji = "普"
result_siong_ji = siong_ji_tsa_siann_bu(siong_ji)
assert result_siong_ji["siann_bu"] == "滂"
assert result_siong_ji["tai_lo"] == "ph"
assert result_siong_ji["tshing_lo"] == "次清"
print(f"\n查反切上字：{siong_ji}")
print(f"siong_ji_tsa_siann_bu 測試結果：{result_siong_ji}")

# 測試 e_ji_tsa_un_bu 函數
ha_ji = "荅"
result_ha_ji = ha_ji_tsa_un_bu(ha_ji)
assert result_ha_ji["un_bu"] == "合"
assert result_ha_ji["tai_lo"] == "ap"
print(f"\n查反切下字：{ha_ji}")
print(f"ha_ji_tsa_un_bu 測試結果：{result_ha_ji}")