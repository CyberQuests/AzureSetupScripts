# 安裝繁體中文（台灣）語言包的 PowerShell 腳本

# 設置錯誤操作偏好，遇到錯誤時停止執行
$ErrorActionPreference = "Stop"

try {
    # 安裝 NuGet 包提供程序（如果尚未安裝）
    if (!(Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Host "正在安裝 NuGet 包提供程序..."
        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
    }

    # 安裝 LanguagePackManagement 模組（如果尚未安裝）
    if (!(Get-Module -ListAvailable -Name LanguagePackManagement)) {
        Write-Host "正在安裝 LanguagePackManagement 模組..."
        Install-Module -Name LanguagePackManagement -Force -Scope CurrentUser
    }

    # 導入 LanguagePackManagement 模組
    Import-Module LanguagePackManagement

    # 安裝繁體中文（台灣）語言包
    Write-Host "正在安裝繁體中文（台灣）語言包..."
    Install-Language -Language "zh-TW" -CopyToSettings

    # 設置 Windows 用戶界面語言
    Write-Host "設置 Windows 用戶界面語言為繁體中文（台灣）..."
    Set-WinUILanguageOverride -Language "zh-TW"

    # 設置 Windows 系統區域設置
    Write-Host "設置 Windows 系統區域設置為繁體中文（台灣）..."
    Set-WinSystemLocale "zh-TW"

    # 設置 Windows 用戶語言列表
    Write-Host "設置 Windows 用戶語言列表..."
    Set-WinUserLanguageList "zh-TW" -Force

    # 設置系統文化特性
    Write-Host "設置系統文化特性為繁體中文（台灣）..."
    Set-Culture zh-TW

    # 複製用戶國際化設置到系統和新用戶
    Write-Host "正在複製用戶國際化設置到系統和新用戶..."
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true

    Write-Host "語言包安裝和設置完成。系統將在 30 秒後重新啟動..."
    Start-Sleep -Seconds 30
    Restart-Computer -Force
}
catch {
    Write-Error "安裝過程中發生錯誤: $_"
    exit 1
}
