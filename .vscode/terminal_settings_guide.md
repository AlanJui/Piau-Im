# Windows Terminal 顯示器位置設定指南

## 目標
讓 Windows Terminal 在 Debug 時自動顯示在右側顯示器（識別號 1）

## 方法 1：修改 Windows Terminal 設定（推薦）

### 步驟 1：開啟 Windows Terminal 設定
1. 開啟 Windows Terminal
2. 按 `Ctrl + Shift + ,` 或點選下拉選單 → 設定
3. 點選左下角的「開啟 JSON 檔案」按鈕

### 步驟 2：加入位置設定
在 JSON 檔案的最外層（與 "profiles" 同一層）加入：

**針對你的顯示器配置（2左+3中+1右）：**
```json
{
    "initialPosition": "4480,50",
    "launchMode": "default",
    "centerOnLaunch": false,
    "windowingBehavior": "useNew",
    "initialRows": 50,
    "initialCols": 100
}
```

### 步驟 3：你的顯示器座標計算

**你的配置：**
- 顯示器 2（左）：1920 x 1080
- 顯示器 3（中）：2560 x 1080
- 顯示器 1（右）：1080 x 1920 ← **直立顯示器**

**計算結果：**
- X 座標 = 1920（左邊寬度）+ 2560（中間寬度）= **4480**
- Y 座標 = 50（離頂部 50 像素）

**為什麼用這些值：**
- `initialRows: 50` - 因為你的右側顯示器是直立的（1920 高度），可以顯示更多列
- `initialCols: 100` - 因為寬度只有 1080，欄數要較少

### 步驟 4：完整設定範例

```json
{
    "$help": "https://aka.ms/terminal-documentation",
    "$schema": "https://aka.ms/terminal-profiles-schema",

    "defaultProfile": "{61c54bbd-c2c6-5271-96e7-009a87ff44bf}",

    // 視窗位置設定（加在這裡）
    "initialPosition": "3840,100",
    "launchMode": "default",
    "centerOnLaunch": false,
    "windowingBehavior": "useNew",

    "profiles": {
        "defaults": {},
        "list": [
            // ... 你的 profiles
        ]
    },

    "schemes": [],
    "actions": []
}
```

## 方法 2：簡單記憶位置法（最容易）

1. **第一次設定：**
   - 啟動 Windows Terminal
   - 手動拖曳到右側顯示器
   - 調整到你喜歡的大小（不要最大化）
   - **正常關閉**視窗（點 X 關閉）

2. **在 settings.json 中確保這些設定：**
   ```json
   {
       "centerOnLaunch": false,
       "windowingBehavior": "useExisting"
   }
   ```

3. 之後開啟時會自動回到上次的位置

## 方法 3：使用 PowerToys FancyZones

如果你有安裝 PowerToys：

1. 開啟 PowerToys
2. 啟用 FancyZones
3. 設定「將新建的視窗移到上次使用的區域」
4. 編輯右側顯示器的區域配置
5. 第一次手動將 Terminal 拖到區域中
6. 下次會自動定位

## 在 VS Code 中使用

你的 launch.json 已經有正確的配置：

```json
{
    "name": "Python: Debug with External Console",
    "type": "debugpy",
    "request": "launch",
    "program": "${file}",
    "console": "externalTerminal",
    "cwd": "${workspaceFolder}",
    "env": {
        "PYTHONUNBUFFERED": "1"
    }
}
```

**使用方式：**
1. 按 F5 開始 Debug
2. 選擇 "Python: Debug with External Console"
3. 終端機視窗會在外部開啟
4. 第一次手動移到右側顯示器
5. 之後會記住位置

## 疑難排解

### 問題：每次都回到主顯示器
**解決方法：**
1. 檢查 Windows Terminal settings.json 中沒有 `"centerOnLaunch": true`
2. 確保有設定 `"windowingBehavior": "useNew"` 或 `"useExisting"`
3. 不要用工作管理員強制關閉，要正常關閉視窗

### 問題：不確定正確的 X 座標
**找出方法：**
1. 開啟 Windows Terminal
2. 移到右側顯示器
3. 在 PowerShell 中執行：
   ```powershell
   Add-Type -AssemblyName System.Windows.Forms
   $form = [System.Windows.Forms.Form]::new()
   $form.StartPosition = 'Manual'
   $form.Location = [System.Drawing.Point]::new(
       [System.Windows.Forms.Cursor]::Position.X,
       [System.Windows.Forms.Cursor]::Position.Y
   )
   Write-Host "目前位置: X=$($form.Location.X), Y=$($form.Location.Y)"
   ```

### 問題：想要程式結束後保持視窗開啟
**在程式結尾加上：**
```python
if __name__ == "__main__":
    try:
        # ... 你的程式碼
        pass
    finally:
        input("\n按 Enter 鍵關閉視窗...")
```

## 快速測試

在終端機執行：
```powershell
wt -w 0 new-tab --title "測試" powershell -NoExit -Command "Write-Host '這是測試視窗'"
```

然後手動移到右側顯示器並關閉，下次應該會記住位置。
