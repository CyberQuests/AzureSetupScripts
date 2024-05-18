# InstallLanguagePack.ps1
Install-PackageProvider -Name NuGet -Force
Install-Module -Name LanguagePackManagement -Force
Import-Module LanguagePackManagement
Install-Language -Language "zh-TW" -CopyToSettings
Set-WinUILanguageOverride -Language "zh-TW"
Set-WinSystemLocale "zh-TW"
Set-WinUserLanguageList "zh-TW" -Force
Set-Culture zh-TW
Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true

# 安裝 Telegram
Write-Output "開始安裝 Telegram..."
$telegramInstaller = "$env:TEMP\telegram_installer.exe"
Invoke-WebRequest -Uri "https://telegram.org/dl/desktop/win64/tsetup-x64.5.0.1.exe" -OutFile $telegramInstaller
Start-Process -FilePath $telegramInstaller -ArgumentList "/verysilent" -Wait
Write-Output "Telegram 安裝完成"
