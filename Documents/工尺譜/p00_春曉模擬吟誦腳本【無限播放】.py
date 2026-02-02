"""p00_春曉模擬吟誦腳本【無限播放】 V0.1.0

孟浩然《春曉》工尺譜吟誦模擬，以純五音版本呈現。
提示：本程式需在 Windows 系統執行，因使用 winsound 模組發聲。
"""

import time
import winsound

# 純五音頻率定義 (只保留 宮商角徵羽)
notes = {
    "合": 392,  # 羽 (Low So)
    "上": 523,  # 宮 (Do)
    "尺": 587,  # 商 (Re)
    "工": 659,  # 角 (Mi)
    "六": 784,  # 徵 (So)
    "五": 880,  # 羽 (La)
}

# 拍長設定
BAN = 800  # 板
YAN = 400  # 眼
PAUSE = 0.15

# 【純五音版】序列：將原先的 "凡" 改為 "工"，旋律更穩重
score = [
    ("春", "工", YAN),
    ("眠", "工", YAN),
    ("不", "工", YAN),
    ("覺", "尺", YAN),
    ("曉", "上", BAN),
    (None, None, 600),
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


def play_pure_five_tones():
    print("--- 孟浩然《春曉》【純五音版】模擬 ---")
    print("當前模式：純五音 (宮商角徵羽)，無 Fa 與 Ti")
    print("提示：按下 [Ctrl + C] 停止播放\n")

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
                    print()
                    time.sleep(duration / 1000)

            print("\n" + "-" * 30 + "\n")
            count += 1
            time.sleep(2)

    except KeyboardInterrupt:
        print("\n\n已停止播放。")


if __name__ == "__main__":
    play_pure_five_tones()
