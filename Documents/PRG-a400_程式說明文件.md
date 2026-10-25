# a400 程式說明文件

依據【漢字注音】工作表之內容，製作成帶有漢字標音之網頁。


## 作業流程

```mermaid
%%{init: {
  "flowchart": { "useMaxWidth": false },
  "themeVariables": { "fontSize": "18px" }
}}%%
flowchart LR
    A([開始]) --> B[初始化 載入環境變數 與 init_logging]
    B --> C{取得 Excel 作用中活頁簿}
    C -- 否 --> C1[回傳 EXIT_CODE_NO_FILE] --> Z([結束])
    C -- 是 --> D[呼叫 process wb]
    D --> E[tng_sing_bang_iah 產製網頁]
    E --> F{V3 是否有字串}
    F -- 否 --> G[無內容 仍回傳 EXIT_CODE_SUCCESS] --> H[回到主流程]
    F -- 是 --> I[build_web_page 組 HTML 內容]
    I --> J[彙整 env 名稱 組 meta 標籤]
    J --> K[create_html_file 輸出 HTML 檔案]
    K --> H[回到主流程]
    H --> L[finally 儲存檔案 save_as_new_file]
    L --> M{儲存是否成功}
    M -- 否 --> M1[回傳 EXIT_CODE_SAVE_FAILURE] --> Z
    M -- 是 --> N[記錄儲存路徑]
    N --> O([結束 EXIT_CODE_SUCCESS])
```

## 循序圖

```mermaid
%%{init: {
  "sequence": { "useMaxWidth": false, "wrap": true, "actorFontSize": 20, "messageFontSize": 18, "noteFontSize": 16 },
  "themeVariables": { "fontSize": "18px" }
}}%%
sequenceDiagram
    actor User as 使用者
    participant Main as 程式主流程
    participant Excel as Excel 活頁簿
    participant TNG as tng_sing_bang_iah
    participant BUILD as build_web_page
    participant TitleSvc as 標題加註處理
    participant DB as SQLite 資料庫
    participant MOD as 模組函式 han_ji_ca_piau_im
    participant PiauIm as PiauIm 物件
    participant HTML as create_html_file
    participant SAVE as save_as_new_file

    User->>Main: 執行 a400_製作標音網頁
    Main->>Excel: 取得作用中活頁簿
    alt 取得失敗
        Main-->>User: 回傳 EXIT_CODE_NO_FILE
    else 取得成功
        Main->>TNG: 產製網頁
        TNG->>Excel: 讀取 env 名稱與設定
        TNG->>BUILD: 建立 HTML 內容

        BUILD->>Excel: 讀取網頁格式 與 標題字元
        BUILD->>TitleSvc: 為標題加註標音
        TitleSvc->>Excel: 讀取語音類型、漢字庫、標音方法
        TitleSvc->>DB: 連線資料庫
        loop 逐字處理標題
            TitleSvc->>MOD: 查找漢字音標
            MOD-->>TitleSvc: 回傳查找結果
            TitleSvc->>PiauIm: 轉標音 與 組合
        end

        BUILD-->>TNG: 回傳標題 HTML 片段
        TNG->>Excel: 讀取 V3 來源字串

        loop 逐格處理正文
            BUILD->>Excel: 讀取字元
            alt 非漢字或標點
                BUILD-->>BUILD: 記錄標點或空白為 span
            else 漢字
                BUILD->>PiauIm: 需要時轉帶調符為 TLPA 與調號
                BUILD->>PiauIm: 生成 ruby 標籤
            end
        end

        TNG->>HTML: 建立 HTML 檔案
        Main->>SAVE: 儲存活頁簿為新檔
        SAVE-->>Main: 回傳儲存路徑
        Main-->>User: 完成訊息 與 EXIT_CODE_SUCCESS
    end
```

## 支援變調標音格式

依據【漢字注音】工作表，製作帶 `<ruby>` 的網頁時，以`【本調】`/`【變調】`樣式，製作 <rt> 中的【漢字標音】。

### 辭彙解釋

- 變調標音格式：`duan⁵⁻³`
- 本調標音格式：`duan⁵`

- 各種【漢字標音法】之標準/本調/變調標音格式對照表

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


流程：

 1. 先完成**`【漢字注音】工作表`**

|     | D | E | F | G | H | I | J | K | L | M | N | O | P | Q | R |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 11  |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |
| 12 | tuan3 | hong5 | ze3 | u2  |   | zing7 | ciu1 | khong1 |   | san1 | liam2 | siu1 | bi5  | sin1 | liok8 |
| 13 | 斷     | 虹     | 霽   | 雨   | ， | 淨     | 秋    | 空      | ， | 山    | 染     | 修    | 眉    | 新    | 綠     |
| 14 | duàn | hóng | zè | wǔ |   | zîng | cīu | kōng  |   | sān | liǎm | sīu | bbí | sīn | liók |

 2. 將`【漢字注音】工作表`中的`【漢字儲存格】`，把`【底色】`改塗成`淺灰色`（#E8E8E8）。以此方式，指示此漢字的【漢字標音】，若`【使用變調標音格式】`被設定為`是`時，則 <rt> 中的【漢字標音】需採 `變調標音格式`；其餘未標示之儲存格則用`本調標音格式`。。

    如：`D13`、`F13`、`I13`、`J13`、`M13`、`O13`、`Q13`。

 3. 在`【env】工作表`，將`【使用變調標音格式】`儲存格，設定成`是`。

    此後可想像，執行 a400.py 程式後，在 `<rt>...</rt>` 之中的`【漢字標音】`內容：
    - 使用**變調標音格式**：`D14`、`F14`、`I14`、`J14`、`M14`、`O14`、`Q14`；
    - 使用**本調標音格式**：`E14`、`G14`、`K14`、`N14`、`P14`。

|     | D | E | F | G | H | I | J | K | L | M | N | O | P | Q | R |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 11  |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |
| 12 | tuan3 | hong5 | ze3 | u2  |   | zing7 | ciu1 | khong1 |   | san1 | liam2 | siu1 | bi5  | sin1 | liok8 |
| 13 | 斷     | 虹     | 霽   | 雨   | ， | 淨     | 秋    | 空      | ， | 山    | 染     | 修    | 眉    | 新    | 綠     |
| 14 |duan⁵⁻³|hong²|ze⁵⁻³|wu³|，|zing⁶⁻⁵|ciu¹⁻⁶|kong¹|，|san¹⁻⁶ |liam³|siu¹⁻⁶|bbi²|sin¹⁻⁶|liok⁸|

## 使用變調標音格式之程式處理邏輯

以下借用【虛擬碼】之概念，透過 PlantUML 之【流程圖腳本】進行程式之處理邏輯描述。

```plantuml
start
repeat
    :設定【漢字儲存格指標】;

    '製作 <ruby> 中的 <rt> 內容
    if (標音方式=上 or 標音方式=上及右) then (yes)
        if (使用變調標音格式) then (yes)
            if (漢字儲存格底色= '#E8E8E8') then (yes)
                :將 ruby tag 的 data-pian-tiau 屬性設定【true】;
                :將【台語音標】依【上邊標音】設定，轉換成【採用變調格式】的【漢字標音】;
                :依【變調格式的漢字標音】製成<rt>內容;
            else (no)
                :將 ruby tag 的 data-pian-tiau 屬性設定【false】;
                :將【台語音標】依【上邊標音】設定，轉換成【採用本調格式】的【漢字標音】;
                :依【本調格式的漢字標音】製成<rt>內容;
            endif
        else (<color:red>no)
            :延用原處理邏輯，製成<rt>內容;
        endif
    endif

    '製作 <ruby> 中的 <rtc> 內容
    if (標音方式=右 or 標音方式=上及右) then (yes)
        :製作 <rt> 內容;
    endif

    :依<rt>及<rtc>內容，製作<ruby>內容;

    ' 製作 <html> 內容
    :將<ruby>內容加入<html>內容;

    :將【漢字儲存格指標】往後移動一格;
repeat while (漢字儲存格尚未讀到【終止文章】？) is (<color:red>yes)
->no;
'程式終止前處理作業
:延用原處理邏輯;
stop
@enduml
```