# 使用者需求

## 反切查拚音功能

### 功能描述

用途：可以在終端機，使用 Python Code 查詢漢字的反切拚音。 

指令格式：

```bash
py a300_反切查拼音.py [參數1：查詢漢字] [參數2：反切拼]
```

參數：

1. 查詢漢字： 1 個中文字
2. 反切拼音： 2 個中文字
   2.1 反切上字：反切拼音參數的第 1 個中文字
   2.2 反切下字：反切拼音參數的第 2 個中文字

### 執行流程

1. 在終端機輸入指令，如：

```bash
py a300_反切查拼音.py 東 德紅
```

2. a300_反切查拼音.py 的程式碼處理輸入的 2 個參數，完成後輸出如下格式資料：

```bash
欲查詢拼音之漢字：[參數1：查詢漢字]

反切拼音為：[參數2：反切拼]

反切上字為：[反切上字]
反切下字為：[反切下字]


## 反切查漢字讀音

1. 利用上字：
   (1) 查聲母的台羅拼音字母；
   (2) 分清濁音：上字的台羅拼音聲調，若是1-4為清；5-8為濁

2. 利用下字：
   (1) 查韻母的台羅拼音字母；
   (2) 辨平/上/去/入聲：下字的台羅拼音聲調，若是1/5為平；2/6為上；3/7為去；4/8為入

3. 利用「清/濁」和「平/上/去/入」，查四聲八調的調號
   | 　 | 平 | 上 | 去 | 入 |
   |----+----+----+----+----|
   | 清 |  1 |  2 |  3 |  4 |
   | 濁 |  5 |  6 |  7 |  8 |

    1: 清平 
    2: 清上
    3: 清去
    4: 清入
    5: 濁平
    6: 濁上
    7: 濁去
    8: 濁入

4. 台羅拚音 = 聲母 + 韻母 + 調號

漢字= "東"

上字= "德" --> 台羅拚音：tik4  --> 聲母 = "t"   --> 調號 = 4 --> 清音
下字= "紅" --> 台羅拚音：hong5 --> 聲母 = "ong" --> 調號 = 5 --> 平聲
由清音+平聲 --> 調號 = 5

台羅拼音 = t + ong + 5 = tong5


