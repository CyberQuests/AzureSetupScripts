# 確保腳本在錯誤時停止
$ErrorActionPreference = "Stop"

# 設定語言為台灣繁體中文
try {
    # 檢查並載入所需要的模組
    if (-not (Get-Module -ListAvailable -Name International)) {
        Write-Output "International 模組不可用，嘗試安裝"
        if (Get-WindowsFeature -Name "International-Core" -ErrorAction SilentlyContinue) {
            Install-WindowsFeature -Name "International-Core" -IncludeAllSubFeature
        } else {
            Write-Error "International 模組不可用且無法安裝"
            exit 1
        }
    }

    Import-Module International -ErrorAction Stop

    Write-Output "International 模組已載入"

    # 安裝語言套件
    Install-Language -Language zh-TW -CopyToSettings -ErrorAction Stop
    Write-Output "語言套件已安裝"

    # 設定系統語言和區域
    Set-WinSystemLocale -SystemLocale zh-TW -ErrorAction Stop
    Set-WinUserLanguageList -LanguageList zh-TW -Force -ErrorAction Stop
    Set-WinUILanguageOverride -Language zh-TW -ErrorAction Stop
    Set-Culture zh-TW -ErrorAction Stop
    Set-WinHomeLocation -GeoId 245 -ErrorAction Stop

    Write-Output "語言設置成功"
} catch {
    Write-Error "設定語言失敗：$_"
}

# 設定安裝路徑
$originalProgressPreference = $ProgressPreference
$ProgressPreference = "SilentlyContinue"

# Helper function to download and install applications
function Install-Application {
    param (
        [string]$Uri,
        [string]$OutFile,
        [string]$Arguments
    )

    try {
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -ErrorAction Stop
        if (Test-Path $OutFile) {
            Start-Process -FilePath $OutFile -ArgumentList $Arguments -Wait
            Remove-Item $OutFile -ErrorAction SilentlyContinue
            Write-Output "應用程式安裝成功: $OutFile"
        } else {
            Write-Error "下載失敗：$OutFile"
        }
    } catch {
        Write-Error "安裝失敗：$_"
    }
}

# 安裝 Google Chrome
Install-Application -Uri "https://dl.google.com/chrome/install/375.126/chrome_installer.exe" -OutFile "$env:TEMP\chrome_installer.exe" -Arguments "/silent /install"

# 安裝 Telegram
Install-Application -Uri "https://telegram.org/dl/desktop/win64/tsetup-x64.5.0.1.exe" -OutFile "$env:TEMP\telegram_installer.exe" -Arguments "/S"

# 安裝 Skype
Install-Application -Uri "https://go.skype.com/windows.desktop.download" -OutFile "$env:TEMP\skype_installer.exe" -Arguments "/quiet"

# 安裝 WhatsApp
Install-Application -Uri "https://web.whatsapp.com/desktop/windows/release/x64/WhatsAppSetup.exe" -OutFile "$env:TEMP\whatsapp_installer.exe" -Arguments "/S"

# 安裝 Visual Studio Code
Install-Application -Uri "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user" -OutFile "$env:TEMP\vscode_installer.exe" -Arguments "/silent /mergetasks=!runcode"

# 安裝 Sandboxie
Install-Application -Uri "https://github.com/sandboxie-plus/Sandboxie/releases/download/v1.13.7/Sandboxie-Classic-x64-v5.68.7.exe" -OutFile "$env:TEMP\sandboxie_installer.exe" -Arguments "/S"

# 恢復進度條設置
$ProgressPreference = $originalProgressPreference
