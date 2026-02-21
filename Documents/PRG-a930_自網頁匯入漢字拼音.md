# a930 自網頁匯入漢字拼音

自帶有 <ruby> / <rt> / <crt> 等 tag 的【網頁檔】（HTML檔案格式）
匯出內容，存入 Excel 檔中的【工作表】。

```html
...
<section class="page cover-page title-page-bg">
  <img
    src="./assets/img/cover.png"
    alt="前赤壁賦封面"
    class="cover-img title-bg-img"
  />
  <div class="title-page">
    <h1 class="title">
      <p>
        《
        <ruby><rb>前</rb><rp>(</rp><rt>堅五曾</rt><rp>)</rp></ruby>
        <ruby><rb>赤</rb><rp>(</rp><rt>經四出</rt><rp>)</rp></ruby>
        <ruby><rb>壁</rb><rp>(</rp><rt>經四邊</rt><rp>)</rp></ruby>
        <ruby><rb>賦</rb><rp>(</rp><rt>艍三喜</rt><rp>)</rp></ruby>
        》
      </p>
    </h1>
    <hr class="cover-divider" />
    <h2 class="author">
      <ruby><rb>北</rb><rp>(</rp><rt>公四邊</rt><rp>)</rp></ruby>
      <ruby><rb>宋</rb><rp>(</rp><rt>公三時</rt><rp>)</rp></ruby>
      ：
      <ruby><rb>蘇</rb><rp>(</rp><rt>沽一時</rt><rp>)</rp></ruby>
      <ruby><rb>軾</rb><rp>(</rp><rt>經四時</rt><rp>)</rp></ruby>
    </h2>
  </div>
</section>
...
```

## Excel 工作表

| 漢字 | 漢字標音 |
| ---- | -------- |
| 《   |          |
| 前   | 堅五曾   |
| 赤   | 經四出   |
| 壁   | 經四邊   |
| 賦   | 艍三喜   |
| 》   |          |

## 程式語言與套件庫

- 程式語言： Python 3
- Python 套件：使用 xlwings

## 程式功能與邏輯：

1. HTML 解析：使用 BeautifulSoup 解析 HTML。
1. 鎖定範圍：自動搜尋 div.title-page（標題區）與 div.content-box（內文區），若找不到則回退搜尋所有段落。
1. 資料提取：

- Ruby 標籤：提取 <rb>（漢字）與 <rt>（標音）。
- 純文字/標點：對於不在 <ruby> 內的文字（如標點符號 《、》、： 等），會逐字提取，並將標音欄位留空。
- 忽略空白：自動過濾排版用的換行與空白字元。
  1.Excel 匯出：使用 xlwings 自動開啟 Excel，建立名為「網頁匯入」的工作表，並將資料填入 A（漢字）、B（漢字標音）兩欄。

## 使用方式：

### 確保已安裝套件：

```powershell
pip install beautifulsoup4 xlwings
```

### 程式執行

執行腳本（若不带參數，預設讀取 《前赤壁賦》.html）：

```powershell-interactive
python a930_自網頁匯入漢字拼音.py
```

**或者指定 HTML 檔案路徑：**

```powershell-interactive
python a930_自網頁匯入漢字拼音.py 路徑/你的檔案.html
```

您能利用我已有的程式，完成【漢字標音】轉換【台語音標】的工作嗎？

前：堅五曾 = 韻 + 調 + 聲 = [ian] + [5] + [z]

1. C 欄：【台語音標】，如： zian5
2. D 槾：【台語音標之聲】，如：z
3. E 槾：【台語音標之韻】，如：ian
4. F 槾：【台語音標之聲】，如：5
