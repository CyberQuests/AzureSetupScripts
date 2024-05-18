# InstallLanguagePack.ps1
Install-PackageProvider -Name NuGet -Force
Install-Module -Name LanguagePackManagement -Force
Import-Module LanguagePackManagement
Install-Language -Language "zh-TW"
Set-WinUILanguageOverride -Language "zh-TW"
Set-WinSystemLocale "zh-TW"
Set-WinUserLanguageList "zh-TW" -Force
