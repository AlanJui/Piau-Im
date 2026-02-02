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
    (None, None, 600),  # 句間停頓稍微加長
    ("處", "六", YAN),
    ("處", "五", YAN),
    ("聞", "六", YAN),
    ("啼", "工", YAN),
    ("鳥", "尺", BAN),
    (None, None, 600),
    ("夜", "六", YAN),
    ("來", "五", YAN),
    ("風", "工", YAN),
    ("雨", "尺", YAN),
    ("聲", "上", BAN),
    (None, None, 600),
    ("花", "尺", YAN),
    ("落", "工", YAN),
    ("知", "尺", YAN),
    ("多", "上", YAN),
    ("少", "合", BAN),
]


def play_gongche_loop():
    print("--- 孟浩然《春曉》工尺譜吟誦模擬 ---")
    print("狀態：無限循環播放中...")
    print("提示：欲結束程式，請按下 [Ctrl + C]\n")

    try:
        count = 1
        while True:
            print(f"【第 {count} 遍】")
            for char, note, duration in score:
                if char:
                    print(f"{char}({note})", end=" ", flush=True)
                    winsound.Beep(notes[note], duration)
                    time.sleep(PAUSE)
                else:
                    print()  # 換行
                    time.sleep(duration / 1000)

            print("\n" + "=" * 30 + "\n")
            count += 1
            time.sleep(2)  # 每一遍結束後停頓 2 秒再開始

    except KeyboardInterrupt:
        print("\n\n偵測到結束指令 (Ctrl+C)。")
        print("吟誦已停止！")


if __name__ == "__main__":
    play_gongche_loop()
