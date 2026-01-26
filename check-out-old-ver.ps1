<#
.SYNOPSIS
    從指定的 Git Commit 取出特定檔案，並儲存到指定的新位置（可更名）。

.DESCRIPTION
    此腳本使用 `git show` 指令從 Git 歷史記錄中提取檔案。
    它解決了 PowerShell 在處理二進位檔案（如 Excel .xlsx）重新導向時可能發生的編碼損毀問題，
    確保檔案完整地被還原。

.PARAMETER CommitId
    Git Commit 的雜湊值 (Hash) 或參照 (如 HEAD, tag)。

.PARAMETER SourceFilePath
    Repo 中的原始檔案路徑（相對路徑）。

.PARAMETER DestinationPath
    輸出的檔案完整路徑（包含目錄與新檔名）。

.EXAMPLE
    .\check-out-old-ver.ps1 -CommitId "529e01a" -SourceFilePath "output7/file.xlsx" -DestinationPath "_tmp/restored.xlsx"
    .\check-out-old-ver.ps1 529e01af8ff91a3f755f78b0d41a7eb73c7b6739 "output7/【河洛文讀注音-閩拼調符】《深慮論》.xlsx" "_tmp/【河洛文讀注音-閩拼調符】《深慮論》_restored.xlsx"
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$CommitId,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$SourceFilePath,

    [Parameter(Mandatory = $true, Position = 2)]
    [string]$DestinationPath
)

Set-StrictMode -Version Latest

# 1. 檢查 Git 是否存在
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "錯誤: 系統找不到 git 指令。"
    exit 1
}

# 2. 處理目的路徑
try {
    # 取得絕對路徑，方便後續處理
    # 如果路徑不存在，Resolve-Path 會失敗，所以先處理目錄
    $DestDir = Split-Path -Path $DestinationPath -Parent

    # 如果目錄是空的（例如只給檔名），預設為當前目錄
    if ([string]::IsNullOrWhiteSpace($DestDir)) {
        $DestDir = "."
    }

    # 如果目標目錄不存在，則建立
    if (-not (Test-Path -Path $DestDir)) {
        Write-Host "建立目錄: $DestDir"
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }

    # 組合最終輸出路徑 (解決相對路徑問題)
    $AbsDestPath = $DestinationPath
    if (-not [System.IO.Path]::IsPathRooted($DestinationPath)) {
        $AbsDestPath = Join-Path (Get-Location) $DestinationPath
    }
}
catch {
    Write-Error "路徑處理失敗: $_"
    exit 1
}

# 3. 執行 Git Show 提取檔案
# 注意：Git 路徑標準使用 forward slash (/)
$GitSourcePath = $SourceFilePath -replace '\\', '/'

Write-Host "正在從 Commit '$CommitId' 提取檔案..."
Write-Host "來源: $GitSourcePath"
Write-Host "目標: $AbsDestPath"

# 4. 使用 cmd /c 執行重新導向
# 說明：PowerShell 5.1 的 '>' 運算符號預設會影響二進位檔案（如 .xlsx）的編碼，導致檔案損毀。
# 使用 cmd /c 直接由 shell 處理 stdout 重新導向是處理二進位檔案最安全簡便的方法。

$CmdCommand = "git show ${CommitId}:`"${GitSourcePath}`" > `"${AbsDestPath}`""

# 執行
cmd /c $CmdCommand

# 5. 檢查結果
if ($LASTEXITCODE -eq 0) {
    if (Test-Path $AbsDestPath) {
        $FileSize = (Get-Item $AbsDestPath).Length
        if ($FileSize -gt 0) {
            Write-Host "✅ 成功！檔案已儲存至: $AbsDestPath ($FileSize bytes)" -ForegroundColor Green
        }
        else {
            Write-Warning "⚠️ 檔案已建立，但是是空的 (0 bytes)。請確認 Commit ID 或檔案路徑是否正確。"
        }
    }
    else {
        Write-Error "❌ 未知錯誤：檔案未被建立。"
    }
}
else {
    Write-Error "❌ Git 指令執行失敗。請確認下列事項："
    Write-Error "   1. Commit ID '$CommitId' 是否正確？"
    Write-Error "   2. 檔案路徑 '$GitSourcePath' 在該 Commit 中是否存在？"
    Write-Error "   (git exit code: $LASTEXITCODE)"
}
