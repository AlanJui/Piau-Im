# 專案指引

## 開發環境建置指引

### 安裝專案所需 Python 套件

```bash
cd <ProjectRootDir>
python -m venv .venv
.venv/Scripts/activate
pip install -r requirements.txt
```

### 安裝 Chrome Web Drvier

本程式之執行，需配合 Chrome WebDriver ，目前預設 Windows 10 安裝路徑：

d:\bin\chromedriver-win64\chromedriver.exe

上述安裝位置，須於 `config.env` 執行環境設定檔指明：

```sh
CHROMEDRIVER_PATH=d:\bin\chromedriver-win64\chromedriver.exe
```
【安裝步驟】：

1. 確認作業系統中安裝的 Google Chrome 瀏覽器版本。方法為：打開瀏覽器輸入如下指令...

```powershell-interactive
chrome://version
```

2. 根據所查到的 Chrome 版本，自 [ChromeDriver 下載頁面](https://chromedriver.chromium.org/home) 下載對應的 ChromeDriver。
   以 Chrome V122.0.6261.95 言，自此[網址](https://googlechromelabs.github.io/chrome-for-testing/#stable)，在表格中的 Binary 欄位，找 chromedriver ，對映的 Platform : win64 下載：[chromedirver-win64.zip](https://storage.googleapis.com/chrome-for-testing-public/122.0.6261.94/win64/chromedriver-win64.zip)

3. 將上述下載之 zip 壓縮檔解開，置於路徑： d:\bin\chromedriver-win64\
