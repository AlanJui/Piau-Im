
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


#--------------------------------------------------------------------------
# 【帶調符音標】轉【帶調符TLPA音標】
#--------------------------------------------------------------------------
def tiau_hu_im_piau_tng_uann(im_piau: str, po_ci: bool = True) -> str:
    #---------------------------------------------------------
    # 保留【音標】之首字母
    #---------------------------------------------------------
    su_ji = ""      # 預設傳入之音標首字母不為大寫
    if po_ci and im_piau[0].isupper():
        su_ji = im_piau[0]
    im_piau = im_piau.lower()
    #---------------------------------------------------------
    # 轉換音標中【聲母】
    #---------------------------------------------------------
    if im_piau.startswith("tsh"):
        im_piau = im_piau.replace("tsh", "c", 1)
    elif im_piau.startswith("chh"):
        im_piau = im_piau.replace("chh", "c", 1)
    elif im_piau.startswith("ts"):
        im_piau = im_piau.replace("ts", "z", 1)
    elif im_piau.startswith("ch"):
        im_piau = im_piau.replace("ch", "z", 1)
    # 如若傳入之【音標】首字母為大寫，則將已轉成 "z" 或 "c" 之拼音字母改為大寫
    if su_ji and im_piau[0] == "c":
        im_piau = "C" + im_piau[1:]
    elif su_ji and im_piau[0] == "z":
        im_piau = "Z" + im_piau[1:]

    #---------------------------------------------------------
    # 轉換音標中【韻母】
    #---------------------------------------------------------
    # un_bu_i_tng_uann = tua_tiau_hu_un_bu_tng_uann(im_piau)
    im_piau_i_tng_uann = tng_un_bu(im_piau)

    return im_piau_i_tng_uann
