# 程式說明文件 a401 Generate Web Page

根據【漢字注音】工作表轉製漢字標音【網頁】。

## 程式架構

1. 配置類別 WebPageConfig

    - 集中管理所有 Excel 配置參數（列數、欄數、標音方式等）
    - 初始化時自動從 Excel Named Range 讀取設定

2. 網頁生成器類別 WebPageGenerator

    - 專職負責 Ruby 標籤生成
    - 處理標題製作
    - 產生完整網頁內容
    - 分離邏輯關切點，便於測試和維護

3. 結構化函數組織

    - generate_web_page() - 主要處理函數
    - process() - 作業流程控制
    - main() - 程式入口

4. 改善的錯誤處理與日誌

    - 完整的 Try-Except-Finally 結構
    - 明確的結束代碼定義
    - 詳細的處理步驟紀錄

5. 程式風格一致性

    - 採用 a210 的註解區隔方式
    - 相同的常數定義格式
    - 一致的函數簽名和文件字串

## 循序圖

1. 程式執行流程循序圖 - 展示完整執行流程

    - main() → process() → generate_web_page()
    - 配置初始化流程
    - 標題處理
    - 內容循環處理
    - 檔案儲存

2. Ruby 標籤生成循序圖 - 詳細的標音邏輯

    - 台語音標分解
    - 聲母判斷
    - 根據網頁格式選擇標音位置
    - Ruby 標籤生成策略

3. 類別互動關係圖 - 展示物件協作

    - WebPageConfig - 配置管理
    - PiauIm - 標音轉換
    - WebPageGenerator - 核心生成邏輯
    - 各函數的角色定位

4. 資料流向圖 - 從輸入到輸出

    - Excel 資料讀取
    - 配置流轉
    - 標音轉換
    - HTML 生成和輸出

```mermaid
sequenceDiagram
    participant User as 使用者
    participant Main as main()
    participant Process as process()
    participant GenWebPage as generate_web_page()
    participant Config as WebPageConfig
    participant PiauIm as PiauIm
    participant Generator as WebPageGenerator
    participant Sheet as Excel Sheet
    participant HTML as HTML 輸出

    User->>Main: 執行程式
    activate Main

    Main->>Main: 初始化日誌
    Main->>Main: 取得程式資訊
    Main->>Main: 記錄程式開始

    Main->>Main: 取得作用中 Excel 活頁簿
    alt 無法取得活頁簿
        Main->>Main: 返回 EXIT_CODE_NO_FILE
        Main->>User: 結束
    end

    Main->>Process: 呼叫 process(wb)
    activate Process

    Process->>Process: 記錄作業開始
    Process->>GenWebPage: 呼叫 generate_web_page(wb)
    activate GenWebPage

    GenWebPage->>Config: 建立 WebPageConfig(wb)
    activate Config
    Config->>Sheet: 讀取 Excel Named Range
    Config->>Sheet: 讀取標音相關參數
    Config->>Sheet: 讀取網頁格式設定
    Config-->>GenWebPage: 返回配置物件
    deactivate Config

    GenWebPage->>PiauIm: 建立 PiauIm(han_ji_khoo_name)
    activate PiauIm
    PiauIm-->>GenWebPage: 返回標音物件
    deactivate PiauIm

    GenWebPage->>Generator: 建立 WebPageGenerator(config, piau_im)
    activate Generator
    Generator-->>GenWebPage: 返回生成器物件
    deactivate Generator

    GenWebPage->>Sheet: 啟用工作表

    GenWebPage->>Generator: 呼叫 generate_web_page(sheet, wb)
    activate Generator

    Generator->>Generator: 生成圖片 HTML Tag
    Generator->>Generator: 呼叫 generate_title_with_ruby()

    loop 讀取標題字元
        Generator->>Sheet: 讀取標題儲存格
        Generator->>Generator: 判斷是否為漢字
        alt 是漢字
            Generator->>Generator: 呼叫 generate_ruby_tag()
            Generator->>Generator: 生成 Ruby 標籤
        else 非漢字
            Generator->>Generator: 直接輸出
        end
    end

    Generator->>Generator: 生成文章 Div 標籤

    loop 逐列處理工作表內容
        Generator->>Sheet: 讀取儲存格內容

        alt 結尾標示 'φ'
            Generator->>Generator: 設置 End_Of_File = True
            Generator->>Generator: 中斷迴圈
        else 換行標示 '\n'
            Generator->>Generator: 輸出換行標籤
            Generator->>Generator: 重置字數計數
        else 非漢字
            alt 標點符號
                Generator->>Generator: 輸出標點符號
            else 空白
                Generator->>Generator: 輸出全形空白
            else 其他字元
                Generator->>Generator: 輸出其他字元
            end
            Generator->>Generator: 字數計數 +1
        else 漢字
            Generator->>Sheet: 讀取漢字標音儲存格
            Generator->>Generator: 轉換為 TLPA 音標
            Generator->>Generator: 呼叫 generate_ruby_tag()
            Generator->>Generator: 生成 Ruby 標籤
            Generator->>Generator: 字數計數 +1
        end

        alt 達到每列字數限制
            Generator->>Generator: 輸出換行標籤
            Generator->>Generator: 重置字數計數
        end
    end

    Generator->>Generator: 輸出結束標籤
    Generator-->>GenWebPage: 返回 HTML 內容
    deactivate Generator

    GenWebPage->>GenWebPage: 生成輸出檔案名稱
    GenWebPage->>GenWebPage: 構建 meta 標籤
    GenWebPage->>HTML: 呼叫 _create_html_file()
    activate HTML
    HTML->>HTML: 組合 HTML 模板
    HTML->>HTML: 建立輸出目錄
    HTML->>HTML: 寫入 HTML 檔案
    HTML-->>GenWebPage: 返回
    deactivate HTML

    GenWebPage->>GenWebPage: 記錄完成訊息
    GenWebPage-->>Process: 返回 EXIT_CODE_SUCCESS
    deactivate GenWebPage

    Process->>Sheet: 啟用【漢字注音】工作表
    Process->>Process: 記錄作業結束
    Process-->>Main: 返回 EXIT_CODE_SUCCESS
    deactivate Process

    Main->>Sheet: 啟用【漢字注音】工作表
    Main->>Main: 儲存檔案 (save_as_new_file)

    alt 儲存成功
        Main->>Main: 記錄儲存路徑
    else 儲存失敗
        Main->>Main: 記錄錯誤訊息
    end

    Main->>Main: 記錄程式結束
    Main->>User: 返回結束代碼
    deactivate Main

```

## WebPageGenerator 內部流程 - Ruby 標籤生成

```mermaid
sequenceDiagram
    participant Generator as WebPageGenerator
    participant PiauIm as PiauIm
    participant RubyTag as Ruby標籤

    Generator->>Generator: generate_ruby_tag(han_ji, tai_gi_im_piau)

    Generator->>Generator: 分解台語音標
    Generator->>Generator: split_tai_gi_im_piau()
    Generator->>Generator: 取得聲母、韻母、聲調

    alt 聲母為空
        Generator->>Generator: 設置聲母為 "ø"
    else 聲母存在
        Generator->>Generator: 使用讀取的聲母
    end

    alt 網頁格式為「無預設」
        alt 標音方式為「上及右」
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回上方標音
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回右方標音
        else 標音方式為「上」
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回上方標音
        else 標音方式為「右」
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回右方標音
        end
    else 網頁格式為特定類型
        alt 格式為 POJ/TL/BP/TLPA_Plus/SNI
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回上方標音
        else 格式為 TPS
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回右方標音
        else 格式為 DBL
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回上方標音
            Generator->>PiauIm: 呼叫 han_ji_piau_im_tng_huan()
            PiauIm-->>Generator: 返回右方標音
        end
    end

    Generator->>Generator: _build_ruby_tag(han_ji, siong_piau_im, zian_piau_im)

    alt 僅上方標音
        Generator->>RubyTag: 生成上方標籤 &lt;rt&gt;
    else 僅右方標音
        Generator->>RubyTag: 生成右方標籤 &lt;rtc&gt;
    else 上方及右方標音
        Generator->>RubyTag: 生成上方 &lt;rt&gt; + 右方 &lt;rtc&gt;
    else 無標音
        Generator->>RubyTag: 生成純文字 &lt;span&gt;
    end

    RubyTag-->>Generator: 返回 Ruby HTML 標籤

```

## 類別互動關係圖

```mermaid
classDiagram
    class WebPageConfig {
        - wb: Workbook
        - TOTAL_LINES: int
        - CHARS_PER_ROW: int
        - ROWS_PER_LINE: int
        - start_row: int
        - start_col: int
        - end_row: int
        - end_col: int
        - han_ji_khoo_name: str
        - ue_im_lui_piat: str
        - piau_im_huat: str
        - piau_im_format: str
        - piau_im_hong_sik: str
        - siong_pinn_piau_im: str
        - zian_pinn_piau_im: str
        - title: str
        - image_url: str
        - output_dir: str
        - zu_im_huat_list: dict
        __init__(wb)
    }

    class PiauIm {
        - han_ji_khoo_name: str
        han_ji_piau_im_tng_huan()
        hong_im_tng_tai_gi_im_piau()
    }

    class WebPageGenerator {
        - config: WebPageConfig
        - piau_im: PiauIm
        __init__(config, piau_im)
        + generate_ruby_tag(han_ji, tai_gi_im_piau)
        + generate_title_with_ruby(sheet, wb)
        + generate_web_page(sheet, wb)
        - _build_ruby_tag(han_ji, siong_piau_im, zian_piau_im)
    }

    class main {
        + main()
    }

    class process {
        + process(wb)
    }

    class generate_web_page {
        + generate_web_page(wb, sheet_name)
        - _create_html_file(output_path, content, title, head_extra)
    }

    main --> process: 呼叫
    process --> generate_web_page: 呼叫
    generate_web_page --> WebPageConfig: 建立
    generate_web_page --> PiauIm: 建立
    generate_web_page --> WebPageGenerator: 建立
    WebPageGenerator --> WebPageConfig: 使用
    WebPageGenerator --> PiauIm: 使用
    WebPageGenerator --> main: 生成HTML

```

## 資料流向圖

```mermaid
graph TD
    A["Excel 活頁簿"] -->|讀取 Named Range| B["WebPageConfig"]
    A -->|讀取工作表| C["WebPageGenerator"]

    B -->|提供配置| C

    D["PiauIm"] -->|標音轉換| C

    C -->|處理標題| E["Ruby 標籤"]
    C -->|處理內容| E
    C -->|處理圖片| E

    E -->|組合| F["HTML 內容"]
    F -->|輸出| G["HTML 檔案"]

    C -->|記錄| H["Console 日誌"]
    B -->|記錄| H

```

