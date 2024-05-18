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


Invoke-WebRequest -Uri "https://telegram.org/dl/desktop/win64/" -OutFile "telegram_install.exe"
Start-Process "telegram_install.exe" -ArgumentList "/S" -Wait

Restart-Computer
