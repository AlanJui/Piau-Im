# 支援變調標音樣式變更規範

## 摘要

本文件用於規範：新增之`支援【標示變調/本調】之【漢字標音】`功能，需進行之變更。

## 文章網頁頂端操作介面

每篇【漢字標音】網頁，其【網頁頂端】之【操作介面】，網頁組件及版面結構，規劃如下：

```html
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>每篇文章頂端的操作介面</title>
    <link rel="stylesheet" href="./assets/styles/styles.css">
</head>

<body>
    <!-- 系統預設【漢字】之【頂端】及【右端】使用之【漢字標音法】 -->
    <label for="phonetic-styles-predefine">預設標音樣式：</label>
    <select id="phonetic-styles-predefine" name="phonetic-styles-predefine">
        <option value="十五音+方音符號">十五音+方音符號</option>
        <option value="台語音標+雅俗通">台語音標+雅俗通</option>
        <option value="台語音標+十五音">台語音標+十五音</option>
        <option value="台語音標+方音符號">台語音標+方音符號</option>
        <option value="閩拼調號+方音符號">閩拼調號+方音符號</option>
        <option value="方音符號">方音符號</option>
        <option value="台語音標">台語音標</option>
        <option value="台羅拼音">台羅拼音</option>
        <option value="白話字">白話字</option>
        <option value="閩拼方案">閩拼方案</option>
    </select>
    <!-- 群組分隔線 -->
    <span
        style="display: inline-block; width: 1px; height: 20px; background-color: #ccc; margin: 0 15px; vertical-align: middle;"></span>

    <!-- 使用者自設【漢字】之【頂端】及【右端】使用之【漢字標音法】 -->
    <label for="phonetic-styles-custom">自訂標音樣式&nbsp;</label>
    <!-- <ruby>/<rt> 適用 -->
    <label for="phonetic-styles-custom-top">上端：</label>
    <select id="phonetic-styles-custom-top" name="phonetic-styles-custome-top">
        <option value="無">無</option>
        <option value="雅俗通">雅俗通</option>
        <option value="十五音">十五音</option>
        <option value="方音符號">方音符號</option>
        <option value="國際音標">國際音標</option>
        <option value="台語音標">台語音標</option>
        <option value="台羅拼音">台羅拼音</option>
        <option value="白話字">白話字</option>
        <option value="閩拼方案">閩拼方案</option>
        <option value="閩拼調號">閩拼調號</option>
        <option value="台語注音二式">台語注音二式</option>
    </select>
    <!-- <ruby>/<rtc> 適用 -->
    <label for="phonetic-styles-custom-right">右端：</label>
    <select id="phonetic-styles-custom-right" name="phonetic-styles-custom-right">
        <option value="無">無</option>
        <option value="雅俗通">雅俗通</option>
        <option value="十五音">十五音</option>
        <option value="方音符號">方音符號</option>
    </select>
    <!-- 群組分隔線 -->
    <span
        style="display: inline-block; width: 1px; height: 20px; background-color: #ccc; margin: 0 15px; vertical-align: middle;"></span>
    <!-- 是否採用本調/變調標音樣式 -->
    <label for="pian-tiau-switcher">變調標音：</label>
    <select id="pian-tiau-switcher" name="pian-tiau-switcher">
        <option value="off">關閉</option>
        <option value="on">啟用</option>
    </select>
</body>

</html>
```

## 變調標音

【變調標音】下拉式選項清單（pian-tiau-switcher），**預設**為：`關閉`，一旦使用者變更為：`啟用`。JavaScript 負責依據`預設標音樣式`或`自訂標音樣式`的`上端`及`右端`為 `<ruby>` 中的 `<rt>` 與/或 `<rtc>` 轉換`【漢字標音】`。

### 變調標音格式

- 變調標音格式：`duan⁵⁻³`
- 本調標音格式：`duan⁵`

各種【漢字標音法】之格式對照表

| 漢字標音法 | 標準標音格式 | 本調標音格式 | 變調標音格式  |
|-------|--------|--------|---------|
| 十五音   | 地觀三    | 地觀三    | 地觀三-七   |
| 雅俗通   | 觀三地    | 地觀三    | 地觀三-七   |
| 方音符號  | ㄉㄨㄢ˪   | ㄉㄨㄢ³   | ㄉㄨㄢ³⁻⁷  |
| 台語音標  | tuan3  | tuan³  | tuan³⁻⁷ |
| 台羅拼音  | tuàn   | tuan³  | tuan³⁻⁷ |
| 白話字   | toàn   | toan³  | toan³⁻⁷ |
| 注音二式  | duan3  | duan³  | duan³⁻⁷ |
| 閩拼調號  | duan5  | duan⁵  | duan⁵⁻³ |
| 閩拼調符  | duàn   | duan⁵  | duan⁵⁻³ |

### 變調標音需 html 支援處

JavaScript 執行【標示變調/本調】的【漢字標音】時，需要自 `<ruby>` 讀取以下兩項`【屬性】`：

- data-tlpa：讀取【漢字】的`【台語音標】`，以為【漢字標音】轉換之依據；
- data-piau-tiau：
  - "true": 表示 <ruby> 中的【漢字】，其`<rt>`中【漢字標音】採用`【變調標音格式】`；
  - "false"：表示 <ruby> 中的【漢字】，其`<rt>`中【漢字標音】採用`【本調標音格式】`；

```html
                ......
				<p>
					<ruby data-tlpa="tuan3" data-pian-tiau="true">斷<rt>tuan³⁻²</rt><rtc>ㄉㄨㄢ˪</rtc></ruby>
					<ruby data-tlpa="hong5" data-pian-tiau="false">虹<rt>hong⁵</rt><rtc>ㄏㆲˊ</rtc></ruby>
					<ruby data-tlpa="ze3" data-pian-tiau="true">霽<rt>ze³⁻²</rt><rtc>ㄗㆤ˪</rtc></ruby>
					<ruby data-tlpa="u2" data-pian-tiau="false">雨<rt>u²</rt><rtc>ㄨˋ</rtc></ruby>
					<span>，</span>
					<ruby data-tlpa="zing7" data-pian-tiau="true">淨<rt>zing⁷⁻³</rt><rtc>ㄐㄧㄥ˫</rtc></ruby>
                    ......
				</p>
```

若是 `<ruby>` 中無 `data-pian-tiau` 屬性，等同 `data-pian-tiau="false"` 。故 `<rt>` 中的【漢字標音】，亦採用`【本調標音格式】`。