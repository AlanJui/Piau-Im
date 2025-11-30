from mod_TL_tiau_hu_tng_tiau_ho import tiau_hu_tng_tiau_ho

test_cases = ["tēng", "tīng", "súi"]

for im_piau in test_cases:
    bo_tiau_im_piau, tiau_ho = tiau_hu_tng_tiau_ho(im_piau)
    print(f"{im_piau} => [{bo_tiau_im_piau}, {tiau_ho}]")
