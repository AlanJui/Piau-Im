# 定義基本拼音字母（韻母包含 m, n）
base_letters = "aeioumn"

# 定義聲調符號
accent_marks = "\u0300\u0301\u0302\u0304\u0306\u030C\u030D"

# 生成所有拼音字母的聲調變體
pinyin_variants = "".join(f"{ch}{accent}" for ch in base_letters for accent in accent_marks)

# 動態生成正規表示式
regex_pattern = rf"[{pinyin_variants}]+$"

# 顯示結果
print(regex_pattern)

# [àáâāăǎa̍èéêēĕěe̍ìíîīĭǐi̍òóôōŏǒo̍ùúûūŭǔu̍m̀ḿm̂m̄m̆m̌m̍ǹńn̂n̄n̆ňn̍]+$
