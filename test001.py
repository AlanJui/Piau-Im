import re


def split_hong_im_hu_ho(hong_im_hu_ho):
    # 定義調符對應的字典
    Hong_Im_Tiau_Hu_Dict = {
        "ˋ": 2,
        "˪": 3,
        "ˊ": 5,
        "˫": 7,
        "\u02D9": 8,  # '˙'
    }

    # 編譯調符的正則表達式模式
    HongImTiauHu = re.compile(r"[ˋ˪ˊ˫˙]", re.I)

    # 定義表示第四聲的尾字元集合
    tone_4_endings = {'ㆴ', 'ㆵ', 'ㆻ', 'ㆷ'}

    # 定義聲母的集合
    sheng_mu_ji = {
        'ㄅ', 'ㄆ', 'ㆠ', 'ㄇ',
        'ㄉ', 'ㄊ', 'ㄋ', 'ㄌ',
        'ㄍ', 'ㄎ', 'ㆣ', 'ㄏ', 'ㄫ',
        'ㄗ', 'ㄘ', 'ㆡ', 'ㄙ',
        'ㄐ', 'ㄑ', 'ㆢ', 'ㄒ',
        'ㄓ', 'ㄔ', 'ㄕ', 'ㄖ',
        'ㄭ', 'ㄪ', 'ㄬ', 'ㄈ',
    }

    # 步驟一：檢查最後一個字元是否為調符
    if HongImTiauHu.match(hong_im_hu_ho[-1]):
        tiau_fu = hong_im_hu_ho[-1]
        tiau_hao = Hong_Im_Tiau_Hu_Dict[tiau_fu]
        # 移除調符，獲得無調符的方音符號
        wu_tiau_fu_hong_im_hu_ho = hong_im_hu_ho[:-1]
    else:
        # 最後沒有調符，判斷是第一聲還是第四聲
        if hong_im_hu_ho[-1] in tone_4_endings:
            tiau_hao = 4
        else:
            tiau_hao = 1
        wu_tiau_fu_hong_im_hu_ho = hong_im_hu_ho

    # 步驟四：提取聲母和韻母
    if wu_tiau_fu_hong_im_hu_ho and wu_tiau_fu_hong_im_hu_ho[0] in sheng_mu_ji:
        sheng_mu = wu_tiau_fu_hong_im_hu_ho[0]
        yun_mu = wu_tiau_fu_hong_im_hu_ho[1:]
    else:
        sheng_mu = ''
        yun_mu = wu_tiau_fu_hong_im_hu_ho

    return [sheng_mu, yun_mu, str(tiau_hao)]

if __name__ == '__main__':
    examples = {
        "ㄍㄨㄧ": "歸",
        "ㄧˋ": "己",
        "ㄍㄧㄥ˪": "敬",
        "ㄍㄚㆻ": "覺",
        "ㄌㄞˊ": "來",
        "ㄒㄧ˫": "是",
        "ㄉㄚㆻ˙": "獨",
    }

    for hong_im_hu_ho, word in examples.items():
        sheng_mu, yun_mu, tiau_hao = split_hong_im_hu_ho(hong_im_hu_ho)
        print(f"詞語：{word}")
        print(f"方音符號：{hong_im_hu_ho}")
        print(f"聲母：{sheng_mu}")
        print(f"韻母：{yun_mu}")
        print(f"調號：{tiau_hao}")
        print("-" * 20)

