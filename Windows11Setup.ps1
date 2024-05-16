# 確保腳本在錯誤時停止
$ErrorActionPreference = "Stop"

# 設定語言為台灣繁體中文
Import-Module International -ErrorAction Stop

Write-Output "International 模組已載入"

# 安裝語言套件
Install-Language -Language zh-TW -CopyToSettings -ErrorAction Stop
Write-Output "語言套件已安裝"

# 設定 Windows 顯示語言為繁體中文（台灣）
Set-WinUILanguageOverride -Language zh-TW -ErrorAction Stop

# 設定系統語言和區域
Set-WinUserLanguageList -LanguageList zh-TW -Force -ErrorAction Stop
Set-WinSystemLocale -SystemLocale zh-TW -ErrorAction Stop
Set-WinHomeLocation -GeoId 208 -ErrorAction Stop

# 設定系統文化為繁體中文（台灣）
Set-Culture zh-TW -ErrorAction Stop

# 安裝 Google Chrome
$chromeInstaller = "$env:TEMP\chrome_installer.exe"
Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/latest/chrome_installer.exe" -OutFile $chromeInstaller
Start-Process -FilePath $chromeInstaller -ArgumentList "/silent /install" -Wait

# 安裝 Telegram
$telegramInstaller = "$env:TEMP\telegram_installer.exe"
Invoke-WebRequest -Uri "https://telegram.org/dl/desktop/win64/tsetup-x64.5.0.1.exe" -OutFile $telegramInstaller
Start-Process -FilePath $telegramInstaller -ArgumentList "/silent /install" -Wait

# 安裝 Skype
$skypeInstaller = "$env:TEMP\skype_installer.exe"
Invoke-WebRequest -Uri "https://go.skype.com/windows.desktop.download" -OutFile $skypeInstaller
Start-Process -FilePath $skypeInstaller -ArgumentList "/silent /install" -Wait

# 安裝 Visual Studio Code
$vscodeInstaller = "$env:TEMP\vscode_installer.exe"
Invoke-WebRequest -Uri "https://update.code.visualstudio.com/latest/win32-x64-user/stable" -OutFile $vscodeInstaller
Start-Process -FilePath $vscodeInstaller -ArgumentList "/silent /mergetasks=!runcode" -Wait

# 安裝 Sandboxie
$sandboxieInstaller = "$env:TEMP\sandboxie_installer.exe"
Invoke-WebRequest -Uri "https://github.com/sandboxie-plus/Sandboxie/releases/download/v1.13.7/Sandboxie-Classic-x64-v5.68.7.exe" -OutFile $sandboxieInstaller
Start-Process -FilePath $sandboxieInstaller -ArgumentList "/verysilent" -Wait

