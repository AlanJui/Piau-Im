# 自動製作打字練習表 - 使用說明

## 功能說明

這個程式可以自動從【漢字注音】工作表中讀取資料，並製作【打字練習表】供打字練習使用。

## 檔案說明

1. `auto_typing_practice.py` - 主程式
2. `test_typing_practice.py` - 測試程式
3. `README_typing_practice.md` - 本說明文件

## 使用前準備

1. 確保已安裝 Python 和 xlwings 套件

   ```bash
   pip install xlwings
   ```

2. 準備 Excel 檔案：
   - 開啟包含【漢字注音】工作表的 Excel 檔案
   - 確保【漢字注音】工作表的格式符合規格

## 漢字注音工作表格式

### 儲存格結構

- 使用範圍：{D3:R2002}
- 每4個儲存格構成一個單元
- 單元結構：
  ```
  第1格(row1): 人工標音欄
  第2格(row2): 台語音標欄
  第3格(row3): 漢字欄        ← 會被複製到打字練習表
  第4格(row4): 漢字標音欄    ← 會被複製到打字練習表
  ```

### 範例
| 位置 | 內容 |
|------|------|
| D3 | tik8 (人工標音) |
| D4 | tik8 (台語音標) |
| D5 | 笛 (漢字) |
| D6 | ㄉㄧㆻ˙ (漢字標音) |

## 使用方法

### 方法1：直接執行
```bash
python auto_typing_practice.py
```

### 方法2：在 Python 中呼叫
```python
from auto_typing_practice import create_typing_practice_sheet
create_typing_practice_sheet()
```

### 方法3：在 Jupyter Notebook 中使用
```python
import sys
sys.path.append('C:/work/Piau-Im')  # 調整為實際路徑

from auto_typing_practice import create_typing_practice_sheet
create_typing_practice_sheet()
```

## 打字練習表格式

### 欄位說明
- B欄：漢字
- C欄：注音符號/羅馬拼音
- E~M欄：分解後的字元（最多9個字元）

### 範例輸出
| B | C | D | E | F | G | H | I | J | K | L | M |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 漢字 | 注音/拼音 |  | 字元1 | 字元2 | 字元3 | 字元4 | 字元5 | 字元6 | 字元7 | 字元8 | 字元9 |
| 笛 | ㄉㄧㆻ˙ |  | ㄉ | ㄧ | ㆻ | (空白) |  |  |  |  |  |
| 同 | tong5 |  | t | o | n | g | / |  |  |  |  |

## 聲調按鍵對照表

### 羅馬拼音聲調
| 聲調 | 調名 | 按鍵 | 範例 |
|------|------|------|------|
| 1 | 陰平 | ; | tong1 → tong; |
| 2 | 陰上 | \ | tong2 → tong\ |
| 3 | 陰去 | _ | tong3 → tong_ |
| 4 | 陰入 | [ | tok4 → tok[ |
| 5 | 陽平 | / | tong5 → tong/ |
| 7 | 陽去 | - | tong7 → tong- |
| 8 | 陽入 | ] | tok8 → tok] |

### 注音符號聲調
| 聲調符號 | 調名 | 按鍵 | 範例 |
|----------|------|------|------|
| (無) | 陰平 | 空白 | ㄙㄨ → ㄙㄨ(空白) |
| ˊ | 陽平 | 6 | ㄌㄧㄤˊ → ㄌㄧㄤ6 |
| ˇ | 陰去 | 3 | ㄌㄧㄤˇ → ㄌㄧㄤ3 |
| ˋ | 陰上 | 4 | ㄌㄧㄤˋ → ㄌㄧㄤ4 |
| ¯ | 陽去 | 5 | ㄌㄧㄤ¯ → ㄌㄧㄤ5 |
| ˙ | 輕聲 | 空白 | ㄉㄧㆻ˙ → ㄉㄧㆻ(空白) |

## 特殊處理

### 入聲調處理
羅馬拼音的入聲調（4、8調）會自動轉換：
- ng結尾 → k結尾（tong4 → tok4）
- n結尾 → t結尾（an4 → at4）
- m結尾 → p結尾（am4 → ap4）

### 終結處理
當遇到漢字 'φ' 時會停止處理。

## 錯誤排除

### 常見問題
1. **ModuleNotFoundError: No module named 'xlwings'**
   - 解決：`pip install xlwings`

2. **找不到【漢字注音】工作表**
   - 確保工作表名稱正確
   - 確保 Excel 檔案已開啟

3. **程式執行沒有反應**
   - 確保 Excel 應用程式在背景執行
   - 確保有作用中的活頁簿

### 偵錯模式
在程式中會輸出處理進度，可以觀察執行狀況。

## 測試

執行測試程式來驗證功能：
```bash
python test_typing_practice.py
```

測試內容包括：
- 聲調對照表
- 拼音分解功能
- 各種輸入格式的處理