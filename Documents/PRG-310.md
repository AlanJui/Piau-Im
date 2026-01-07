# a310_缺字表修正後續作業 程式架構說明

## 摘要

這份文檔包含：

📋 程式概述與功能說明
🏗️ 完整的架構層級圖
📦 各層級與函數的詳細說明
🔄 執行流程與資料流向圖
🎯 設計原則與最佳實踐
📊 函數職責判斷準則
🛠️ 使用方式與退出碼
🔍 關鍵函數深入說明
📚 相關模組列表
🎓 學習要點
🔄 與其他程式的關係

## 📋 程式概述

**程式名稱**：`a310_缺字表修正後續作業.py`

**主要功能**：
1. 讀取【缺字表】工作表中已補標音的漢字資料
2. 將校正後的音標更新回【漢字注音】工作表
3. 將校正音標轉換為台羅拼音並寫入 SQLite 資料庫
4. 更新【標音字庫】工作表的資料記錄

**使用時機**：當【缺字表】中的漢字已補填音標後，執行後續的整合與同步作業

---

## 🏗️ 程式架構層級

### 完整架構圖

```
架構層級：
┌─────────────────────────────────────────────────────────┐
│ 主程式入口（main）                                        │
│ └── 程式流程控制、錯誤處理、檔案儲存                        │
└─────────────────────────────────────────────────────────┘
                         ↓ 調用
┌─────────────────────────────────────────────────────────┐
│ 業務邏輯層（Business Logic Layer）                        │
│ ├── process()                  - 整體作業流程協調          │
│ ├── update_khuat_ji_piau()     - 批次處理缺字表            │
│ └── khuat_ji_piau_poo_im_piau() - 批次寫入資料庫          │
└─────────────────────────────────────────────────────────┘
         ↓ 調用                    ↓ 調用
┌────────────────────┐    ┌──────────────────────────────┐
│ 基礎設施層/輔助層    │    │ 資料處理層                    │
│ （Helper Functions）│    │                              │
│                    │    │ ┌──────────────────────────┐ │
│ _initialize_ji_khoo()   │ │ ProcessConfig (類別)      │ │
│ └── 初始化作業      │    │ │ └── 配置參數管理          │ │
│     資料準備        │    │ └──────────────────────────┘ │
│                    │    │                              │
│ _save_ji_khoo_to_excel()│ │ ┌──────────────────────────┐ │
│ └── 資料持久化      │    │ │ CellProcessor (類別)      │ │
│     寫回 Excel      │    │ │ └── 單一儲存格處理邏輯    │ │
│                    │    │ └──────────────────────────┘ │
│ _process_sheet()   │    │                              │
│ └── 批次迭代        │    │ ┌──────────────────────────┐ │
│     控制迴圈        │    │ │ DatabaseManager          │ │
│   （本程式未使用）  │    │ │ └── 資料庫連線管理        │ │
│                    │    │ └──────────────────────────┘ │
│ tiau_zing_piau_im_ │    │                              │
│  ji_khoo_dict()    │    │ ┌──────────────────────────┐ │
│ └── 字典資料整理    │    │ │ JiKhooDict               │ │
│                    │    │ │ └── 字庫資料結構          │ │
└────────────────────┘    │ └──────────────────────────┘ │
                          └──────────────────────────────┘
```

---

## 📦 各層級詳細說明

### 1️⃣ 主程式入口層

#### `main(args) -> int`

**職責**：
- 程式生命週期管理
- 全域錯誤處理
- 檔案儲存與結果輸出

**執行流程**：
```python
def main(args):
    # 1. 取得作用中 Excel 活頁簿
    wb = xw.apps.active.books.active

    # 2. 執行業務邏輯
    exit_code = process(wb, args)

    # 3. 儲存檔案
    save_as_new_file(wb)

    # 4. 返回執行結果
    return exit_code
```

---

### 2️⃣ 業務邏輯層

#### `process(wb, args) -> int`

**職責**：協調整體作業流程

**執行步驟**：
```python
def process(wb, args):
    # 步驟 1：初始化配置與資料
    config = ProcessConfig(wb, args)
    ji_khoo_dicts = _initialize_ji_khoo(wb, ...)
    processor = CellProcessor(config, *ji_khoo_dicts)

    # 步驟 2：處理缺字表（更新 Excel）
    update_khuat_ji_piau(wb, config, processor)

    # 步驟 3：回填資料庫
    khuat_ji_piau_poo_im_piau(wb, config, processor)

    # 步驟 4：儲存字庫
    _save_ji_khoo_to_excel(wb, *ji_khoo_dicts)

    return EXIT_CODE_SUCCESS
```

#### `update_khuat_ji_piau(wb, config, processor) -> int`

**職責**：批次處理缺字表，更新漢字注音工作表

**處理流程**：
1. 讀取【缺字表】中的每一列資料
2. 將【校正音標】轉換為標準 TLPA+ 格式
3. 根據座標欄位，找到【漢字注音】工作表中對應的儲存格
4. 更新【台語音標】儲存格（漢字上方一列）
5. 更新【漢字標音】儲存格（漢字下方一列）
6. 清除【漢字】儲存格的底色標記
7. 更新【標音字庫】的記憶體資料結構

**輸入**：
- `wb`: Excel 活頁簿物件
- `config`: 配置參數物件
- `processor`: 處理器物件（包含字庫資料）

**輸出**：執行狀態碼（EXIT_CODE_SUCCESS 或錯誤碼）

#### `khuat_ji_piau_poo_im_piau(wb, config, processor) -> int`

**職責**：批次將缺字表資料寫入 SQLite 資料庫

**處理流程**：
1. 讀取【缺字表】工作表的所有資料
2. 將 TLPA+ 音標轉換為台羅拼音（TL）
3. 使用 `DatabaseManager` 執行資料庫操作：
   - 檢查【漢字 + 音標】組合是否已存在
   - 若存在則更新，不存在則新增
4. 根據語音類型設定常用度（文讀音=0.8，白話音=0.6）

**輸入**：
- `wb`: Excel 活頁簿物件
- `config`: 配置參數物件
- `processor`: 處理器物件（包含資料庫管理器）

**輸出**：執行狀態碼（EXIT_CODE_SUCCESS 或錯誤碼）

---

### 3️⃣ 基礎設施層/輔助層

> **命名規則**：所有基礎設施層函數以底線 `_` 開頭，表示內部使用

#### `_initialize_ji_khoo(wb, ...) -> tuple[JiKhooDict, ...]`

**職責**：資料準備與初始化

**功能**：
- 從 Excel 工作表讀取既有的字庫資料
- 建立三個 `JiKhooDict` 物件：
  1. 人工標音字庫
  2. 標音字庫
  3. 缺字表
- 可選擇性清空工作表（新建模式）

**類比**：開工前準備工具和材料

**參數說明**：
- `new_jin_kang_piau_im_ji_khoo_sheet`: 是否重建人工標音字庫工作表
- `new_piau_im_ji_khoo_sheet`: 是否重建標音字庫工作表
- `new_khuat_ji_piau_sheet`: 是否重建缺字表工作表

#### `_save_ji_khoo_to_excel(wb, ...)`

**職責**：資料持久化

**功能**：
- 將記憶體中的 `JiKhooDict` 物件寫回 Excel
- 更新三個字庫工作表的內容

**類比**：工作完成後保存成果

**何時調用**：在 `process()` 函數結束前

#### `_process_sheet(sheet, config, processor)`

**職責**：迭代控制器

**功能**：
- 控制迴圈，遍歷工作表中的所有儲存格
- 調用 `CellProcessor.process_cell()` 處理每個儲存格

**類比**：流水線的傳送帶

**注意**：⚠️ **本程式（a310）不需要此函數**，因為 a310 是批次處理整個工作表，而非逐個儲存格處理

#### `tiau_zing_piau_im_ji_khoo_dict(piau_im_ji_khoo_dict, ...)`

**職責**：資料結構維護

**功能**：
- 調整【標音字庫】字典的內部資料結構
- 移除舊座標
- 新增或更新音標資料

**類比**：資料庫的索引維護

---

### 4️⃣ 資料處理層

#### `ProcessConfig` 類別

**職責**：配置參數管理（資料層）

**屬性分類**：

**Excel 工作表結構參數**：
```python
self.TOTAL_LINES          # 每頁總列數
self.ROWS_PER_LINE        # 每行佔用的 Excel 列數（4列）
self.CHARS_PER_ROW        # 每列總字數
self.line_start_row       # 起始列號
self.line_end_row         # 結束列號
self.start_col            # 起始欄號
self.end_col              # 結束欄號
```

**儲存格位置偏移量**：
```python
self.jin_kang_piau_im_row_offset = 0   # 人工標音儲存格偏移
self.tai_gi_im_piau_row_offset = 1     # 台語音標儲存格偏移
self.han_ji_row_offset = 2              # 漢字儲存格偏移
self.han_ji_piau_im_row_offset = 3     # 漢字標音儲存格偏移
```

**業務相關參數**：
```python
self.han_ji_khoo_name     # 漢字庫名稱（河洛話/廣韻）
self.db_name              # 資料庫檔案名稱
self.piau_im_huat         # 標音方法
self.ue_im_lui_piat       # 語音類型（文讀音/白話音）
```

**工具物件**：
```python
self.ji_tian              # HanJiTian 字典查詢物件
self.piau_im              # PiauIm 標音轉換物件
```

#### `CellProcessor` 類別

**職責**：儲存格處理邏輯（操作層）

**主要屬性**：
```python
self.config                      # ProcessConfig 配置物件
self.ji_tian                     # 字典查詢工具
self.piau_im                     # 標音轉換工具
self.db_manager                  # 資料庫管理器
self.jin_kang_piau_im_ji_khoo   # 人工標音字庫
self.piau_im_ji_khoo            # 標音字庫
self.khuat_ji_piau_ji_khoo      # 缺字表字庫
```

**設計特點**：
- 將配置參數、工具物件、字庫資料整合在一起
- 便於在函數間傳遞參數
- 避免函數參數過多

#### `DatabaseManager` 類別

**職責**：資料庫連線管理

**主要功能**：
```python
# 連線管理
db_manager.connect(db_path)      # 建立連線
db_manager.disconnect()          # 關閉連線

# 查詢操作
db_manager.execute(sql, params)  # 執行 SQL
db_manager.fetchone(sql, params) # 查詢單筆
db_manager.fetchall(sql, params) # 查詢多筆

# 交易管理
with db_manager.transaction():
    # 自動 commit 或 rollback
    db_manager.execute(...)
```

**優勢**：
- 單例模式，避免重複建立連線
- 自動管理交易（commit/rollback）
- 統一的錯誤處理
- 程式碼更簡潔

#### `JiKhooDict` 類別

**職責**：字庫資料結構

**資料格式**：
```python
{
    "漢字": [
        {
            "tai_gi_im_piau": "台語音標",
            "kenn_ziann_im_piau": "校正音標",
            "coordinates": [(row1, col1), (row2, col2), ...]
        },
        ...
    ],
    ...
}
```

**主要方法**：
- `add_entry()`: 新增項目
- `update_entry()`: 更新項目
- `get_value_by_key()`: 查詢值
- `write_to_excel_sheet()`: 寫回 Excel
- `create_ji_khoo_dict_from_sheet()`: 從 Excel 讀取

---

## 🔄 完整執行流程

### 主要流程圖

```
┌─────────────┐
│  main()     │ 程式入口
└──────┬──────┘
       │
       ↓
┌─────────────────────────┐
│ process(wb, args)       │ 業務邏輯協調
└────┬────────────────────┘
     │
     ├─→ _initialize_ji_khoo()        ← 初始化字庫
     │   └─→ JiKhooDict.create_...()
     │
     ├─→ ProcessConfig(wb, args)      ← 建立配置
     │
     ├─→ CellProcessor(config, ...)   ← 建立處理器
     │   └─→ DatabaseManager.connect()
     │
     ├─→ update_khuat_ji_piau()       ← 處理缺字表
     │   ├─→ 讀取【缺字表】
     │   ├─→ 轉換音標格式
     │   ├─→ 更新【漢字注音】工作表
     │   └─→ tiau_zing_piau_im_ji_khoo_dict()
     │
     ├─→ khuat_ji_piau_poo_im_piau()  ← 回填資料庫
     │   ├─→ 讀取【缺字表】
     │   ├─→ 轉換 TLPA+ → 台羅拼音
     │   └─→ insert_or_update_to_db()
     │       └─→ DatabaseManager.transaction()
     │
     ├─→ _save_ji_khoo_to_excel()     ← 儲存字庫
     │   └─→ JiKhooDict.write_to_excel_sheet()
     │
     └─→ DatabaseManager.disconnect()  ← 關閉連線
```

### 資料流向圖

```
【缺字表】工作表（Excel）
    │
    ├─→ update_khuat_ji_piau()
    │   └─→ 【漢字注音】工作表（更新音標）
    │       └─→ 【標音字庫】記憶體字典
    │
    └─→ khuat_ji_piau_poo_im_piau()
        └─→ 【漢字庫】資料表（SQLite）
            └─→ insert_or_update_to_db()
                ├─→ 新增記錄（INSERT）
                └─→ 更新記錄（UPDATE）
```

---

## 🎯 設計原則與最佳實踐

### 1. 職責分離原則（Separation of Concerns）

```
每個函數/類別都有明確的單一職責：
✓ ProcessConfig     → 只管配置參數
✓ CellProcessor     → 只處理儲存格
✓ DatabaseManager   → 只管資料庫
✓ 業務邏輯函數      → 只實現業務需求
✓ 輔助函數         → 只提供基礎服務
```

### 2. 依賴注入（Dependency Injection）

```python
# ✅ 好的做法：透過參數傳入依賴
def update_khuat_ji_piau(wb, config: ProcessConfig, processor: CellProcessor):
    db_manager = processor.db_manager  # 使用注入的 db_manager

# ❌ 不好的做法：在函數內部建立依賴
def update_khuat_ji_piau(wb):
    db_manager = DatabaseManager()  # 緊耦合
```

### 3. 使用底線命名內部函數

```python
# 公開函數（給外部使用）
def update_khuat_ji_piau():
    pass

# 內部函數（只在模組內使用）
def _initialize_ji_khoo():
    pass
```

### 4. 錯誤處理策略

```python
# 業務邏輯層：捕捉並記錄錯誤，返回錯誤碼
def process(wb, args):
    try:
        update_khuat_ji_piau(wb, config, processor)
    except Exception as e:
        logging_exc_error("處理缺字表失敗", e)
        return EXIT_CODE_PROCESS_FAILURE

# 主程式層：處理全域錯誤
def main(args):
    try:
        return process(wb, args)
    except Exception as e:
        logging_exc_error("程式異常終止", e)
        return EXIT_CODE_UNKNOWN_ERROR
```

### 5. 資源管理

```python
# ✅ 使用 finally 確保資源釋放
try:
    khuat_ji_piau_poo_im_piau(wb, config, processor)
except Exception as e:
    logging_exc_error("回填資料庫失敗", e)
finally:
    if processor.db_manager:
        processor.db_manager.disconnect()

# ✅ 使用 context manager 自動管理交易
with db_manager.transaction():
    db_manager.execute("INSERT INTO ...")
    # 自動 commit，出錯自動 rollback
```

---

## 📊 函數職責判斷準則

### 如何判斷函數應該放在哪一層？

#### 業務邏輯層函數特徵：
- ✓ 直接實現業務需求（「更新缺字表」、「回填資料庫」）
- ✓ 包含業務規則和邏輯判斷
- ✓ 會被主程式直接調用
- ✓ 函數名稱描述業務動作

#### 基礎設施層函數特徵（`_` 開頭）：
- ✓ 提供基礎服務（初始化、儲存、迭代）
- ✓ 不包含業務邏輯，只做技術性工作
- ✓ 只被業務邏輯層調用，不直接被 `main()` 調用
- ✓ 函數名稱描述技術動作

#### 資料處理層（類別）特徵：
- ✓ 管理資料狀態（ProcessConfig）
- ✓ 封裝單一職責的操作（CellProcessor）
- ✓ 提供可重用的資料處理能力
- ✓ 類別名稱描述資料或處理器

---

## 🛠️ 使用方式

### 命令列參數

```bash
# 一般模式：使用既有字庫
python a310_缺字表修正後續作業.py

# 新建模式：重建所有字庫工作表
python a310_缺字表修正後續作業.py --new

# 測試模式
python a310_缺字表修正後續作業.py --test
```

### 執行前置條件

1. Excel 檔案必須已開啟並為作用中活頁簿
2. 必須包含以下工作表：
   - 【漢字注音】
   - 【缺字表】
   - 【人工標音字庫】
   - 【標音字庫】
3. 【缺字表】中的【校正音標】欄（C欄）已填入音標
4. 資料庫檔案（Ho_Lok_Ue.db 或 Kong_Un.db）存在於專案目錄

---

## 📝 退出碼說明

| 退出碼 | 常數名稱 | 說明 |
|-------|---------|------|
| 0 | EXIT_CODE_SUCCESS | 執行成功 |
| 1 | EXIT_CODE_NO_FILE | 無法找到檔案或活頁簿 |
| 2 | EXIT_CODE_INVALID_INPUT | 輸入資料錯誤 |
| 3 | EXIT_CODE_SAVE_FAILURE | 檔案儲存失敗 |
| 10 | EXIT_CODE_PROCESS_FAILURE | 處理過程失敗 |
| 99 | EXIT_CODE_UNKNOWN_ERROR | 未知錯誤 |

---

## 🔍 關鍵函數說明

### `insert_or_update_to_db()`

**功能**：將漢字與音標資料插入或更新至資料庫

**特點**：
- 使用 `DatabaseManager` 統一管理連線
- 使用 `transaction()` context manager 自動管理交易
- 檢查【漢字 + 音標】組合是否已存在（而非只檢查漢字）
- 根據標音方法設定常用度（文讀音=0.8，白話音=0.6）

**改進說明**：
```python
# ❌ 舊版：只檢查漢字
cursor.execute("SELECT 識別號 FROM 漢字庫 WHERE 漢字 = ?", (han_ji,))

# ✅ 新版：檢查漢字+音標組合
db_manager.fetchone(
    "SELECT 識別號 FROM 漢字庫 WHERE 漢字 = ? AND 台羅音標 = ?",
    (han_ji, tai_gi_im_piau)
)
```

### `tiau_zing_piau_im_ji_khoo_dict()`

**功能**：維護【標音字庫】字典的資料一致性

**處理邏輯**：
1. 在【標音字庫】中搜尋指定的【漢字 + 音標】組合
2. 若找到，從該項目中移除指定的座標
3. 將校正後的音標作為新項目加入字典

**用途**：確保當音標修正後，字庫中的資料能正確更新

---

## 📚 相關模組說明

### 使用的自訂模組

| 模組名稱 | 主要功能 |
|---------|---------|
| `mod_database` | 資料庫連線管理 |
| `mod_ca_ji_tian` | 漢字字典查詢（HanJiTian） |
| `mod_excel_access` | Excel 工作表操作 |
| `mod_字庫` | 字庫資料結構（JiKhooDict） |
| `mod_帶調符音標` | 音標格式轉換 |
| `mod_標音` | 標音轉換與漢字標音 |
| `mod_logging` | 日誌記錄 |

---

## 🎓 學習要點

### 1. 類別 vs 函數的選擇

**何時使用類別**：
- 需要維護狀態（如 ProcessConfig 的配置參數）
- 需要封裝相關的資料和方法（如 CellProcessor）
- 需要單例模式（如 DatabaseManager）

**何時使用函數**：
- 無狀態的操作（如音標轉換）
- 一次性的批次處理（如 update_khuat_ji_piau）
- 簡單的輔助功能（如 _initialize_ji_khoo）

### 2. 參數傳遞的簡化

**問題**：函數需要很多參數時，調用變得複雜
```python
# ❌ 參數過多，難以維護
def update_khuat_ji_piau(wb, db_name, piau_im_huat, ue_im_lui_piat,
                         ji_tian, piau_im, start_row, end_row, ...):
    pass
```

**解決**：使用類別封裝相關參數
```python
# ✅ 參數簡化，易於維護
def update_khuat_ji_piau(wb, config: ProcessConfig, processor: CellProcessor):
    # 所有需要的參數都在 config 和 processor 中
    db_name = config.db_name
    piau_im = processor.piau_im
```

### 3. 資料庫操作的最佳實踐

**使用 DatabaseManager 的好處**：
```python
# ✅ 新版：使用 DatabaseManager
with db_manager.transaction():
    db_manager.execute("INSERT ...")
    # 自動 commit，錯誤時自動 rollback
    # 不需要手動管理連線

# ❌ 舊版：手動管理
conn = sqlite3.connect(db_path)
cursor = conn.cursor()
try:
    cursor.execute("INSERT ...")
    conn.commit()
except:
    conn.rollback()
finally:
    conn.close()
```

---

## 🔄 與其他程式的關係

### 前置程式

- `a200_查找及填入漢字標音.py` - 初次標音，產生缺字表
- `a230_缺字表補標音.py` - 為缺字表補填音標

### 後續程式

- `a330_以標音字庫更新漢字注音工作表.py` - 使用更新後的字庫重新標音
- `a800_更新漢字庫之漢字標音.py` - 批次更新漢字庫

### 工作流程

```
1. a200 初次標音 → 產生【缺字表】
2. a230 人工補標音 → 更新【缺字表】中的【校正音標】欄
3. a310 後續處理 → 將校正音標同步到【漢字注音】與資料庫 ✓ 當前程式
4. a330 重新標音 → 使用更新後的字庫重新標註全文
```

---

## 📅 版本歷史

### v2.0 - 2026/01/06
- ✨ 引入 `DatabaseManager` 統一管理資料庫連線
- ✨ 使用 `ProcessConfig` 和 `CellProcessor` 類別簡化參數傳遞
- 🐛 修正 `insert_or_update_to_db()` 只檢查漢字的問題
- 🐛 修正函數調用時缺少參數的問題
- ♻️ 重構程式架構，明確劃分各層職責
- 📝 完善程式文檔

### v1.0 - 原始版本
- 基本功能實現

---

## 📞 技術支援

如有問題，請參考：
- 專案 README.md
- 各模組的說明文件
- 程式碼內的註解

---

**文件建立日期**：2026/01/06
**最後更新日期**：2026/01/06
**文件版本**：1.0
