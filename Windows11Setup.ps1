# 確保腳本在錯誤時停止
$ErrorActionPreference = "Stop"

# 將腳本設置為 UTF-8 編碼
$scriptPath = "C:\Packages\Plugins\Microsoft.Compute.CustomScriptExtension\1.10.15\Downloads\0\Windows11Setup.ps1"
$content = Get-Content -Path $scriptPath -Raw
Set-Content -Path $scriptPath -Value $content -Encoding UTF8

# 設定語言為台灣繁體中文
Import-Module International -ErrorAction Stop
Write-Output "International 模組已載入"

# 安裝語言套件
Write-Output "開始安裝語言套件..."
Install-Language -Language zh-TW -CopyToSettings -ErrorAction Stop
Write-Output "語言套件已安裝"

# 設定 Windows 顯示語言為繁體中文（台灣）
Write-Output "設定 Windows 顯示語言為繁體中文（台灣）..."
Set-WinUILanguageOverride -Language zh-TW -ErrorAction Stop
Write-Output "Windows 顯示語言設定完成"

# 設定系統語言和區域
Write-Output "設定系統語言和區域..."
Set-WinUserLanguageList -LanguageList zh-TW -Force -ErrorAction Stop
Set-WinSystemLocale -SystemLocale zh-TW -ErrorAction Stop
Set-WinHomeLocation -GeoId 208 -ErrorAction Stop
Write-Output "系統語言和區域設定完成"

# 設定系統文化為繁體中文（台灣）
Write-Output "設定系統文化為繁體中文（台灣）..."
Set-Culture zh-TW -ErrorAction Stop
Write-Output "系統文化設定完成"

# 將設置應用到歡迎屏幕和新用戶
Write-Output "將設置應用到歡迎屏幕和新用戶..."
Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
Write-Output "設置應用完成"

# 重新啟動系統以應用更改
Restart-Computer
