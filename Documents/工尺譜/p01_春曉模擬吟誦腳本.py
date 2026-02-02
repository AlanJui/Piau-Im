import time
import winsound

# 定義頻率 (Hz) - 以 C 大調 (中呂調) 為基準
notes = {
    "合": 392,  # Low So (G4)
    "上": 523,  # Do (C5)
    "尺": 587,  # Re (D5)
    "工": 659,  # Mi (E5)
    "凡": 698,  # Fa (F5)
    "六": 784,  # So (G5)
    "五": 880,  # La (A5)
}

# 定義拍長 (毫秒)
BAN = 800  # 板 (強拍/長音)
YAN = 400  # 眼 (弱拍/短音)
PAUSE = 0.2  # 字與字之間的微小停頓

# 《春曉》工尺譜序列 (字, 工尺譜音, 延時)
score = [
    ("春", "工", YAN),
    ("眠", "凡", YAN),
    ("不", "工", YAN),
    ("覺", "尺", YAN),
    ("曉", "上", BAN),
    (None, None, 400),  # 句間停頓
    ("處", "六", YAN),
    ("處", "五", YAN),
    ("聞", "六", YAN),
    ("啼", "工", YAN),
    ("鳥", "尺", BAN),
    (None, None, 400),
    ("夜", "六", YAN),
    ("來", "五", YAN),
    ("風", "工", YAN),
    ("雨", "尺", YAN),
    ("聲", "上", BAN),
    (None, None, 400),
    ("花", "尺", YAN),
    ("落", "工", YAN),
    ("知", "尺", YAN),
    ("多", "上", YAN),
    ("少", "合", BAN),
]


def play_gongche(sequence):
    print("開始模擬工尺譜吟誦：孟浩然《春曉》\n")
    for char, note, duration in sequence:
        if char:
            print(f"{char} ({note})", end=" ", flush=True)
            # winsound.Beep(頻率, 毫秒)
            winsound.Beep(notes[note], duration)
            time.sleep(PAUSE)
        else:
            print("\n")
            time.sleep(duration / 1000)
    print("\n吟誦結束。")


if __name__ == "__main__":
    play_gongche(score)
