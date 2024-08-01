# Windows11Setup.ps1

# 設置錯誤操作偏好
$ErrorActionPreference = "Stop"

# 函數：安裝應用程式
function Install-Application {
    param (
        [string]$Name,
        [string]$Url,
        [string]$Installer,
        [string]$Arguments
    )
    
    try {
        Write-Output "開始安裝 $Name..."
        $installerPath = "$env:TEMP\$Installer"
        Invoke-WebRequest -Uri $Url -OutFile $installerPath -UseBasicParsing
        Start-Process -FilePath $installerPath -ArgumentList $Arguments -Wait
        Remove-Item -Path $installerPath -Force
        Write-Output "$Name 安裝完成"
    }
    catch {
        Write-Error "安裝 $Name 時發生錯誤: $_"
    }
}

# 檢查磁碟空間
$freeSpace = (Get-PSDrive C).Free / 1GB
if ($freeSpace -lt 5) {
    Write-Error "磁碟空間不足，至少需要 5GB 可用空間。目前可用空間: $($freeSpace)GB"
    exit 1
}

# 安裝應用程式
$applications = @(
    @{
        Name = "Google Chrome"
        Url = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
        Installer = "chrome_installer.exe"
        Arguments = "/silent /install"
    },
    @{
        Name = "Telegram"
        Url = "https://telegram.org/dl/desktop/win64"
        Installer = "telegram_installer.exe"
        Arguments = "/VERYSILENT /NORESTART"
    },
    @{
        Name = "Visual Studio Code"
        Url = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user"
        Installer = "vscode_installer.exe"
        Arguments = "/VERYSILENT /NORESTART /MERGETASKS=!runcode"
    },
    @{
        Name = "Sandboxie"
        Url = "https://github.com/sandboxie-plus/Sandboxie/releases/latest/download/Sandboxie-Classic-x64-v5.68.7.exe"
        Installer = "sandboxie_installer.exe"
        Arguments = "/S"
    }
)

foreach ($app in $applications) {
    Install-Application @app
}

Write-Output "所有應用程式安裝完成"
