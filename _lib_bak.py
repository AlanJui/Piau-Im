
#---------------------------------------------------------------------------------------------=
# 【帶調符韻母轉換】
#---------------------------------------------------------------------------------------------=
def tua_tiau_hu_un_bu_tng_uann(im_piau: str) -> str:
    # 轉換音標中【韻母】為【o͘】（oo長音）的特殊處理
    im_piau = handle_o_dot(im_piau)

    # 轉換音標中【韻母】部份，不含【o͘】（oo長音）的特殊處理
    bo_tiau_hu_im_piau, tone = separate_tone(im_piau)   # 無調符音標：bo_tiau_hu_im_piau
    sorted_keys = sorted(un_bu_mapping, key=len, reverse=True)

    for key in sorted_keys:
        if key in bo_tiau_hu_im_piau:
            bo_tiau_hu_im_piau = bo_tiau_hu_im_piau.replace(key, un_bu_mapping[key])
            break

    # 調符
    # if tone: print(f"調符：{hex(ord(tone))}")
    if tone:
        bo_tiau_hu_im_piau = apply_tone(bo_tiau_hu_im_piau, tone)

    return bo_tiau_hu_im_piau

