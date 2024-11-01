# configure_windows.ps1
# 用於配置 Windows 虛擬機的 PowerShell 腳本

# 設置嚴格的錯誤處理
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 定義日誌函數
function Log-Message {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Severity = 'Info'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Output "[$timestamp] $Severity - $Message"
}

# 檢查網絡連接
function Test-NetworkConnection {
    $testUrl = "http://www.microsoft.com"
    try {
        Invoke-WebRequest -Uri $testUrl -UseBasicParsing -TimeoutSec 10 | Out-Null
        return $true
    } catch {
        Log-Message "無法連接到網絡。錯誤: $_" -Severity 'Error'
        return $false
    }
}

# 設置語言包
function Install-LanguagePack {
    Log-Message "開始安裝繁體中文（台灣）語言包..."
    $Language = 'zh-TW'
    
    try {
        $LangList = Get-WinUserLanguageList
        $Lang = New-WinUserLanguage $Language
        $LangList.Add($Lang)
        Set-WinUserLanguageList $LangList -Force

        # 安裝語言功能
        $languageFeatures = @(
            "Language.Basic~~~$Language~0.0.1.0",
            "Language.Handwriting~~~$Language~0.0.1.0",
            "Language.OCR~~~$Language~0.0.1.0",
            "Language.Speech~~~$Language~0.0.1.0",
            "Language.TextToSpeech~~~$Language~0.0.1.0"
        )

        foreach ($feature in $languageFeatures) {
            Add-WindowsCapability -Online -Name $feature -ErrorAction SilentlyContinue
            if ($?) {
                Log-Message "成功安裝語言功能: $feature"
            } else {
                Log-Message "無法安裝語言功能: $feature" -Severity 'Warning'
            }
        }

        # 設置系統語言
        Set-WinSystemLocale -SystemLocale $Language
        Set-WinUILanguageOverride -Language $Language
        Set-Culture -CultureInfo $Language

        # 設置輸入法
        $languageList = New-WinUserLanguageList -Language $Language
        $languageList[0].InputMethodTips.Clear()
        $languageList[0].InputMethodTips.Add('0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{6024B45F-5C54-11D4-B921-0080C882687E}')
        Set-WinUserLanguageList $languageList -Force

        # 複製語言設置到默認用戶配置
        Copy-Item -Path "$env:USERPROFILE\AppData\Local\Microsoft\Windows\LanguageOverlayCache" -Destination "C:\Users\Default\AppData\Local\Microsoft\Windows\LanguageOverlayCache" -Recurse -Force

        # 複製 UI 語言設置
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\MUI\UILanguages"
        Copy-ItemProperty -Path "$regPath\zh-TW" -Destination "$regPath\en-US" -Recurse -Force

        Log-Message "語言包安裝和設置完成。系統需要重啟以應用更改。"
    } catch {
        Log-Message "安裝語言包時發生錯誤: $_" -Severity 'Error'
        throw
    }
}

# 配置遠程桌面
function Configure-RemoteDesktop {
    Log-Message "開始配置遠程桌面連接..."
    try {
        Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections' -Value 0
        Enable-NetFirewallRule -DisplayGroup 'Remote Desktop'
        Log-Message "遠程桌面連接已啟用，防火牆規則已更新"
    } catch {
        Log-Message "配置遠程桌面時發生錯誤: $_" -Severity 'Error'
        throw
    }
}

# 創建一般用戶
function Create-GeneralUsers {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Users,
        [Parameter(Mandatory=$true)]
        [string]$InitialPassword
    )
    foreach ($user in $Users) {
        try {
            Log-Message "開始創建用戶: $user"
            $password = ConvertTo-SecureString $InitialPassword -AsPlainText -Force
            New-LocalUser -Name $user -Password $password -PasswordNeverExpires:$false
            Add-LocalGroupMember -Group 'Users' -Member $user
            Add-LocalGroupMember -Group 'Remote Desktop Users' -Member $user

            # 設置關機權限
            $ntrightsPath = "C:\Windows\System32\ntrights.exe"
            if (Test-Path $ntrightsPath) {
                $result = Start-Process -FilePath $ntrightsPath -ArgumentList "+r SeShutdownPrivilege -u $user" -Wait -NoNewWindow -PassThru
                if ($result.ExitCode -ne 0) {
                    Log-Message "無法設置用戶 $user 的關機權限" -Severity 'Warning'
                }
            } else {
                Log-Message "警告: ntrights.exe 不存在，無法授予關機權限。請手動設置。" -Severity 'Warning'
            }

            Set-LocalUser -Name $user -PasswordExpired $true

            # 設置用戶語言
            Set-WinUserLanguageList -LanguageList zh-TW -Force -User $user
            Set-Culture -CultureInfo zh-TW -User $user
            Set-WinSystemLocale -SystemLocale zh-TW -User $user
            Set-WinHomeLocation -GeoId 206 -User $user

            $languageList = New-WinUserLanguageList -Language "zh-TW"
            $languageList[0].InputMethodTips.Clear()
            $languageList[0].InputMethodTips.Add('0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{6024B45F-5C54-11D4-B921-0080C882687E}')
            Set-WinUserLanguageList $languageList -Force -User $user

            Log-Message "用戶 $user 創建完成，並已設置語言和權限"
        } catch {
            Log-Message "創建用戶 $user 時發生錯誤: $_" -Severity 'Error'
            throw
        }
    }
}

# 設置基本密碼政策
function Set-PasswordPolicy {
    Log-Message "開始設置密碼策略..."
    try {
        $tempFile = "$env:TEMP\secpol.cfg"
        SecEdit.exe /export /cfg $tempFile
        $secpol = Get-Content $tempFile
        $secpol = $secpol -replace 'MinimumPasswordLength = \d+', 'MinimumPasswordLength = 6'
        $secpol = $secpol -replace 'PasswordComplexity = \d+', 'PasswordComplexity = 0'
        Set-Content $tempFile $secpol
        SecEdit.exe /configure /db secedit.sdb /cfg $tempFile /areas SECURITYPOLICY
        Remove-Item $tempFile -Force
        Log-Message "密碼策略設置完成"
    } catch {
        Log-Message "設置密碼政策時發生錯誤: $_" -Severity 'Error'
        throw
    }
}

# 設置本地組策略
function Set-LocalGroupPolicy {
    Log-Message "開始設置用戶文件夾權限..."
    $users = Get-ChildItem 'C:\Users' -Directory | Where-Object { $_.Name -notin @('Public', 'Default', 'Default User', 'All Users', 'Administrator') }
    foreach ($user in $users) {
        try {
            Log-Message "正在設置文件夾 $($user.FullName) 的權限..."
            icacls "$($user.FullName)" /inheritance:r
            icacls "$($user.FullName)" /grant:r "$($user.Name):(OI)(CI)F" "SYSTEM:(OI)(CI)F" "Administrators:(OI)(CI)F"
            icacls "$($user.FullName)" /deny "Users:(OI)(CI)RX"
            Log-Message "文件夾 $($user.FullName) 權限設置完成"
        } catch {
            Log-Message "設置文件夾 $($user.FullName) 權限時發生錯誤: $_" -Severity 'Error'
            throw
        }
    }
}

# 安裝應用程序
function Install-Applications {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ChromeUrl,
        [Parameter(Mandatory=$true)]
        [string]$TelegramUrl,
        [Parameter(Mandatory=$true)]
        [string]$VSCodeUrl
    )
    
    function Install-Application {
        param (
            [string]$Url,
            [string]$Filename,
            [string]$Arguments,
            [string]$AppName
        )
        
        $tempPath = "C:\Temp"
        if (-not (Test-Path $tempPath)) {
            New-Item -ItemType Directory -Path $tempPath | Out-Null
        }
        
        $installer = Join-Path $tempPath $Filename
        try {
            Log-Message "正在下載 $AppName..."
            Invoke-WebRequest -Uri $Url -OutFile $installer -UseBasicParsing
            
            Log-Message "正在安裝 $AppName..."
            $process = Start-Process -FilePath $installer -ArgumentList $Arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -ne 0) {
                throw "安裝過程返回了非零退出碼: $($process.ExitCode)"
            }
            Log-Message "$AppName 安裝完成"
        } catch {
            Log-Message "安裝 $AppName 時發生錯誤: $_" -Severity 'Error'
            throw
        } finally {
            if (Test-Path $installer) {
                Remove-Item $installer -Force
            }
        }
    }

    if (-not (Test-NetworkConnection)) {
        Log-Message "無法連接到網絡，跳過應用程序安裝" -Severity 'Warning'
        return
    }

    Install-Application -Url $ChromeUrl -Filename "chrome_installer.exe" -Arguments "/silent /install" -AppName "Google Chrome"
    Install-Application -Url $TelegramUrl -Filename "telegram_installer.exe" -Arguments "/VERYSILENT /NORESTART" -AppName "Telegram"
    Install-Application -Url $VSCodeUrl -Filename "vscode_installer.exe" -Arguments "/VERYSILENT /NORESTART /MERGETASKS=!runcode" -AppName "Visual Studio Code"
}

# 主要執行函數
function Main {
    param (
        [Parameter(Mandatory=$true)]
        [array]$GeneralUsers,
        [Parameter(Mandatory=$true)]
        [string]$InitialPassword,
        [Parameter(Mandatory=$true)]
        [string]$ChromeUrl,
        [Parameter(Mandatory=$true)]
        [string]$TelegramUrl,
        [Parameter(Mandatory=$true)]
        [string]$VSCodeUrl
    )

    Log-Message "開始執行 Windows 配置腳本..."

    try {
        Install-LanguagePack
        Configure-RemoteDesktop
        Create-GeneralUsers -Users $GeneralUsers -InitialPassword $InitialPassword
        Set-PasswordPolicy
        Set-LocalGroupPolicy
        Install-Applications -ChromeUrl $ChromeUrl -TelegramUrl $TelegramUrl -VSCodeUrl $VSCodeUrl

        Log-Message "Windows 配置完成。準備執行 Sysprep..."
    
        # 停止 Windows Update 服務
        Stop-Service -Name wuauserv -Force
        # 清除 Windows Update 緩存
        Remove-Item -Recurse -Force C:\Windows\SoftwareDistribution -ErrorAction SilentlyContinue

        # 移除所有用戶的 Appx 包，避免 Sysprep 失敗
        Get-AppxPackage -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue

        # 執行 Sysprep
        $sysprepPath = 'C:\Windows\System32\Sysprep\sysprep.exe'
        if (Test-Path $sysprepPath) {
            Log-Message "開始執行 Sysprep..."
            $sysprepProcess = Start-Process -FilePath $sysprepPath -ArgumentList "/generalize /oobe /shutdown /quiet /mode:vm" -Wait -NoNewWindow -PassThru
            if ($sysprepProcess.ExitCode -ne 0) {
                Log-Message "Sysprep 執行失敗，退出碼: $($sysprepProcess.ExitCode)" -Severity 'Error'
            } else {
                Log-Message "Sysprep 執行完成，虛擬機將關閉。"
            }
        } else {
            Log-Message "找不到 Sysprep 執行檔，無法執行 Sysprep" -Severity 'Error'
        }
    } catch {
        Log-Message "執行過程中發生錯誤: $_" -Severity 'Error'
        throw
    }
}

# 參數驗證
$generalUsers = $env:GENERAL_USERS -split ' '
$initialPassword = $env:INITIAL_PASSWORD
$chromeUrl = $env:CHROME_URL
$telegramUrl = $env:TELEGRAM_URL
$vscodeUrl = $env:VSCODE_URL

# 驗證必要參數
if (-not $generalUsers -or $generalUsers.Count -eq 0) {
    Log-Message "錯誤：未提供一般用戶列表" -Severity 'Error'
    exit 1
}

if ([string]::IsNullOrWhiteSpace($initialPassword)) {
    Log-Message "錯誤：未提供初始密碼" -Severity 'Error'
    exit 1
}

if ([string]::IsNullOrWhiteSpace($chromeUrl)) {
    Log-Message "錯誤：未提供 Chrome 下載 URL" -Severity 'Error'
    exit 1
}

if ([string]::IsNullOrWhiteSpace($telegramUrl)) {
    Log-Message "錯誤：未提供 Telegram 下載 URL" -Severity 'Error'
    exit 1
}

if ([string]::IsNullOrWhiteSpace($vscodeUrl)) {
    Log-Message "錯誤：未提供 Visual Studio Code 下載 URL" -Severity 'Error'
    exit 1
}

# 執行主函數
try {
    Main -GeneralUsers $generalUsers -InitialPassword $initialPassword -ChromeUrl $chromeUrl -TelegramUrl $telegramUrl -VSCodeUrl $vscodeUrl
} catch {
    Log-Message "腳本執行過程中發生嚴重錯誤: $_" -Severity 'Error'
    exit 1
}

Log-Message "Windows 配置腳本執行完畢。"
