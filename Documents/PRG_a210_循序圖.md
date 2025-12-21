# a210 循序圖

為【漢字】自【漢字庫】查找【台語音標】，並以此轉換成【漢字標音】；但在遇有【人工標音】時，則不用在漢字庫查找【台語音標】，而是改以【人工標音】，轉換成【漢字標音】。

## 主要流程：

1. 初始化階段

    - 取得 Excel 活頁簿
    - 讀取配置參數（語音類型、漢字庫）
    - 初始化資料庫連接

2. 準備階段

    - 建立配置物件（ProcessConfig）
    - 初始化字典查詢（HanJiTian）
    - 載入三個字庫（人工標音、標音、缺字表）

3. 處理階段

    - 逐列逐欄掃描 Excel 儲存格
    - 判斷內容類型（人工標音/漢字/特殊字元）
    - 查詢資料庫取得讀音
    - 寫入音標到 Excel

4. 完成階段

    - 儲存三個字庫到 Excel 工作表
    - 回傳處理結果

## 關鍵決策點：

✅ 有人工標音 → 優先使用
✅ 查到讀音 → 寫入標音字庫
❌ 查無讀音 → 記錄到缺字表
🔚 遇到 φ → 結束處理

## 循序圖

```mermaid
sequenceDiagram
    actor User as 使用者
    participant Excel as Excel_VBA
    participant Main as main
    participant CaHanJi as ca_han_ji_thak_im
    participant Config as ProcessConfig
    participant JiTian as HanJiTian
    participant Processor as CellProcessor
    participant Sheet as process_sheet
    participant JiKhoo as JiKhooDict

    User->>Excel: 執行巨集
    Excel->>Main: RunPython

    Main->>Main: 取得活頁簿
    alt 從Excel呼叫
        Main->>Excel: caller
        Excel-->>Main: wb
    else 取得作用中
        Main->>Excel: active
        Excel-->>Main: wb
    end

    Main->>Excel: 讀取語音類型
    Excel-->>Main: ue_im_lui_piat
    Main->>Excel: 讀取漢字庫
    Excel-->>Main: han_ji_khoo

    Main->>CaHanJi: 呼叫處理函數
    activate CaHanJi

    CaHanJi->>Config: 初始化配置
    Config->>Excel: 讀取參數
    Excel-->>Config: 參數
    Config-->>CaHanJi: config

    CaHanJi->>JiTian: 初始化字典
    JiTian-->>CaHanJi: ji_tian

    CaHanJi->>JiKhoo: 初始化人工標音字庫
    JiKhoo-->>CaHanJi: 字庫1
    CaHanJi->>JiKhoo: 初始化標音字庫
    JiKhoo-->>CaHanJi: 字庫2
    CaHanJi->>JiKhoo: 初始化缺字表
    JiKhoo-->>CaHanJi: 字庫3

    CaHanJi->>Processor: 建立處理器
    Processor-->>CaHanJi: processor

    CaHanJi->>Sheet: 處理工作表
    activate Sheet

    loop 每一列
        loop 每一欄
            Sheet->>Excel: 選取儲存格
            Excel-->>Sheet: cell

            Sheet->>Processor: process_cell
            activate Processor

            alt 有人工標音
                Processor->>Excel: 寫入台語音標
                Processor->>Excel: 寫入漢字標音
                Processor->>JiKhoo: 記錄到字庫
            else 漢字
                Processor->>JiTian: 查詢讀音
                JiTian-->>Processor: 結果
                alt 找到
                    Processor->>Excel: 寫入音標
                    Processor->>JiKhoo: 記錄
                else 查無
                    Processor->>JiKhoo: 缺字表
                end
            end

            Processor-->>Sheet: 結果
            deactivate Processor
        end
    end

    Sheet-->>CaHanJi: 完成
    deactivate Sheet

    CaHanJi->>JiKhoo: 寫回人工標音字庫
    CaHanJi->>JiKhoo: 寫回標音字庫
    CaHanJi->>JiKhoo: 寫回缺字表

    CaHanJi-->>Main: SUCCESS
    deactivate CaHanJi

    Main-->>Excel: 回傳
    Excel-->>User: 完成
```