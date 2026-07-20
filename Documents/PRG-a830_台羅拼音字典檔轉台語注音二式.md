# a830 台羅拼音字典檔轉台語注音二式

將 `RIME 字典檔`（以下簡稱：`字典檔`），原先使用之【羅馬字拼音系統】，自【台羅拼音】轉換成【台語注音二式】。

1. 使用【台羅拼音】的原字典檔：**xyz.dict.yaml**，轉換成【台語注音二式】新字典檔，其檔名更換成：**xyz_bpm2.dict.yaml**；

2. 需要轉換的字典檔清單：
   - **【閩南話辭彙】**： ji_khoo_su_lui.dict.yaml    ==》 ji_khoo_su_lui_bpm2.dict.yaml
   - **【泉漳厦閩南字/辭】**： ji_khoo_ban_lam.dict.yaml   ==》 ji_khoo_ban_lam_bpm2.dict.yaml
   - **【閩南話漢語正字】**： ji_khoo_ziann_ji.dict.yaml  ==》 ji_khoo_ziann_ji_bpm2.dict.yaml

3. 【轉換作業】執行時，主要之作業步驟如下所述：
   1. **檔頭**的名稱需更換：每個字典檔的`檔頭`，亦需要更換【name】欄
   2. 轉換工作，在【檔身】中的【台羅拼音】

4. 轉換作業執行環境
   - 來源目錄路徑： C:\Users\AlanJui\work\rime-tlpa\
   - 標的目錄路徑： C:\Users\AlanJui\work\rime-tlpa\

    【舉例】：字典檔 ji_khoo_ziann_ji.dict.yaml
        - 取用時，自路徑 C:\Users\AlanJui\work\rime-tlpa\ji_khoo_ziann_ji.dict.yaml 取用；
        - 存檔時，存到路徑 C:\Users\AlanJui\work\rime-tlpa\ji_khoo_ziann_ji_bpm2.dict.yaml


## 名詞定義

### 字典名稱

參考**字典檔範例**，第 7 行的【name】，便是定義【字典名稱】的所在。以此例而言，其名稱為：`ji_khoo_ziann_ji`。當【台羅拼音】的字典檔，轉換成【台語注音二式】後，需變更名稱為：【原名稱】+【_bpm2】，如：`ji_khoo_ziann_ji_bpm2`

### 字典檔頭

1. 參考**字典檔範例**，第 1-17 行為：`檔案`；
2. 第 17 行為【檔頭與檔身分隔符號】，為三個點字元：【...】。要注意，有些 yaml linter 或 Auto Formatter 會將之視為【語法錯誤】，自動改成其它字元。故需小心，不能令此事發生，否則 RIME 會發生難以預期的異常處理行為。

### 字典檔身

1. 參考**字典檔範例**，第 18-檔尾 行為：`檔身`；
2. 檔身的每筆資料，在第1與第2個 tab 控制字元之間，是為【code】欄位，羅馬拼音字便是存於【code】欄內。

### 【字典檔範例】

```yaml
01 # Rime dictionary
02 # encoding: utf-8
03 #
04 # 閩南語白話音漢字正字
05 #
06 ---
07 name: ji_khoo_ziann_ji
08 version: "v0.1.0"
09 sort: by_weight
10 use_preset_vocabulary: false
11 columns:
12   - text    # 漢字
13   - code    # 台灣音標（TLPA）拼音
14   - weight  # 常用度（優先顯示度）
15   - stem    # 用法舉例
16   - create  # 建立時間
17 ...
18 叫是	kio3 si7	0.60 	以為是	2026-07-05 11:18
19 即例	tsit4 le7	0.60 	這個	2026-07-06 10:59
10 即個	tsit4 e5	0.60 	這個	2026-07-06 10:59
21 ......
```