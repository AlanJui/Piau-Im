# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 語言設定

**所有對話回覆與產出文件，一律使用繁體中文。**

---

## 專案目的

本專案為台語（河洛話）漢字標音系統。主要功能：將 Excel 活頁簿中的漢字，自 SQLite 字典資料庫查得台語讀音，並產出帶有 Ruby 標籤的 HTML 網頁，供網站發布使用。

---

## 開發環境建置

```bash
python -m venv .venv
.venv/Scripts/activate          # Windows
pip install -r requirements.txt
```

部分程式需使用 Selenium，須安裝 ChromeDriver，並於 `config.env` 設定路徑：

```
CHROMEDRIVER_PATH=d:\bin\chromedriver-win64\chromedriver.exe
```

---

## 常用開發指令

```bash
# 程式碼格式化
black .
ruff format .

# 程式碼檢查
ruff check .
pylint .          # 依 .pylintrc 設定，僅啟用 F（嚴重）與 E（錯誤）等級檢查

# 型別檢查
mypy .

# 直接執行單支程式（大多數程式含 __main__ 進入點）
python a200_查找及填入漢字標音.py

# 重新產生 docs/index.html
python a999_自動生成index_html.py
```

程式碼風格：行寬上限 160 字元、字串使用雙引號（設定於 `pyproject.toml` 與 `setup.cfg`）。

---

## 系統架構

### 核心模組層（`mod_*.py`）

| 模組 | 用途 |
|------|------|
| `mod_程式.py` | 程式架構基底：`Program`（設定參數）與 `ExcelCell`（儲存格處理器）類別 |
| `mod_標音.py` | 音標轉換：`PiauIm` 類別，處理 TLPA↔TL↔BP↔MPS2 等系統互轉 |
| `mod_ca_ji_tian.py` | SQLite 查詢：`HanJiTian` 類別負責漢字讀音查找 |
| `mod_字庫.py` | Excel 轉字典：`JiKhooDict` 管理記憶體內字庫 |
| `mod_excel_access.py` | xlwings 工具函式：儲存格定址、工作表 CRUD、命名範圍 |
| `mod_database.py` | SQLite 連線管理員（Singleton 模式），支援 context manager |
| `mod_logging.py` | 日誌工具，輸出至 `process_log.txt` 與 `error_log.txt` |
| `mod_帶調符音標.py` | 調號符號處理工具 |

### 應用程式腳本（`a###_*.py`）

腳本以編號命名，代表作業流程階段：

- **a000–a002**：重置／清除工作表
- **a100**：將文字檔內容填入 Excel 漢字欄位
- **a200–a260**：標音查找與填入（核心標音流程）
- **a300–a320**：人工校正標音
- **a400**：產生含 Ruby 標籤的 HTML 輸出
- **a500–a530**：標音資料匯入／匯出
- **a600–a622**：廣韻（Kong_Un）與十五音字典查找
- **a700–a750**：拚音打字練習工作表產生器
- **a800–a890**：漢字庫維護與匯出
- **a900–a910**：批次處理
- **a999**：產生 `docs/index.html` 供 GitHub Pages 使用

### Excel 工作表儲存格結構

每個漢字欄位依固定列次存放資料：

- **第 1 列**：人工標音（優先覆蓋自動標音）
- **第 2 列**：台語音標（台羅拼音 TL/TLPA）
- **第 3 列**：漢字
- **第 4 列**：漢字標音（依目標音標系統轉換後的最終標音）

### 資料庫

- `Ho_Lok_Ue.db` — 主要河洛話漢字字典
- `Kong_Un.db` — 廣韻（中古漢語）字典
- `雅俗通十五音字典.db` — 十五音字典

資料庫選用由 `.env` 環境變數控制。

### 支援音標系統

TLPA、TL（台羅）、BP（閩拼）、MPS2（方音符號）、POJ（白話字）、十五音。
轉換邏輯集中於 `mod_標音.py` 的 `PiauIm` 類別。

### 輸出目錄

- `docs/` — HTML 檔，經 `.github/workflows/static.yml` 自動部署至 GitHub Pages
- `output*/` — 各次作業的中間與最終 Excel／文字輸出
