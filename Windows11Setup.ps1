# 設定語言為台灣繁體中文
Set-WinUILanguageOverride -Language en-US
Set-WinUserLanguageList -LanguageList zh-TW -Force
Set-WinSystemLocale zh-TW
Set-Culture zh-TW
Set-WinHomeLocation -GeoId 245
Set-WinUILanguage zh-TW

# 設定安裝路徑
$ProgressPreference = "SilentlyContinue"

# 安裝 Google Chrome
Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/375.126/chrome_installer.exe" -OutFile "chrome_installer.exe"
Start-Process -FilePath "chrome_installer.exe" -ArgumentList "/silent /install" -Wait
Remove-Item "chrome_installer.exe"

# 安裝 Telegram
Invoke-WebRequest -Uri "https://telegram.org/dl/desktop/win64/tsetup-x64.5.0.1.exe" -OutFile "telegram_installer.exe"
Start-Process -FilePath "telegram_installer.exe" -ArgumentList "/S" -Wait
Remove-Item "telegram_installer.exe"

# 安裝 Skype
Invoke-WebRequest -Uri "https://go.skype.com/windows.desktop.download" -OutFile "skype_installer.exe"
Start-Process -FilePath "skype_installer.exe" -ArgumentList "/quiet" -Wait
Remove-Item "skype_installer.exe"

# 安裝 Visual Studio Code
Invoke-WebRequest -Uri "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user" -OutFile "vscode_installer.exe"
Start-Process -FilePath "vscode_installer.exe" -ArgumentList "/silent /mergetasks=!runcode" -Wait
Remove-Item "vscode_installer.exe"

# 安裝 Sandboxie
Invoke-WebRequest -Uri "https://github.com/sandboxie-plus/Sandboxie/releases/download/v1.13.7/Sandboxie-Classic-x64-v5.68.7.exe" -OutFile "sandboxie_installer.exe"
Start-Process -FilePath "sandboxie_installer.exe" -ArgumentList "/S" -Wait
Remove-Item "sandboxie_installer.exe"
