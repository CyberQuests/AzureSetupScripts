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

# 安裝 Google Chrome
Write-Output "開始安裝 Google Chrome..."
$chromeInstaller = "$env:TEMP\chrome_installer.exe"
Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/latest/chrome_installer.exe" -OutFile $chromeInstaller
Start-Process -FilePath $chromeInstaller -ArgumentList "/silent /install" -Wait
Write-Output "Google Chrome 安裝完成"

# 安裝 Telegram
Write-Output "開始安裝 Telegram..."
$telegramInstaller = "$env:TEMP\telegram_installer.exe"
Invoke-WebRequest -Uri "https://telegram.org/dl/desktop/win64/tsetup-x64.5.0.1.exe" -OutFile $telegramInstaller
Start-Process -FilePath $telegramInstaller -ArgumentList "/verysilent" -Wait
Write-Output "Telegram 安裝完成"

# 安裝 Visual Studio Code
Write-Output "開始安裝 Visual Studio Code..."
$vscodeInstaller = "$env:TEMP\vscode_installer.exe"
Invoke-WebRequest -Uri "https://update.code.visualstudio.com/latest/win32-x64-user/stable" -OutFile $vscodeInstaller
Start-Process -FilePath $vscodeInstaller -ArgumentList "/verysilent /mergetasks=!runcode" -Wait
Write-Output "Visual Studio Code 安裝完成"

# 安裝 Sandboxie
Write-Output "開始安裝 Sandboxie..."
$sandboxieInstaller = "$env:TEMP\sandboxie_installer.exe"
Invoke-WebRequest -Uri "https://github.com/sandboxie-plus/Sandboxie/releases/download/v1.13.7/Sandboxie-Classic-x64-v5.68.7.exe" -OutFile $sandboxieInstaller
Start-Process -FilePath $sandboxieInstaller -ArgumentList "/verysilent" -Wait
Write-Output "Sandboxie 安裝完成"
