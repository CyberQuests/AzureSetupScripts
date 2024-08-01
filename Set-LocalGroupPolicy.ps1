# Set-LocalGroupPolicy.ps1
# 此腳本用於修改 AppLocker 政策，允許用戶在 Program Files 目錄下安裝和運行應用程式

# 設置嚴格的錯誤處理
$ErrorActionPreference = "Stop"

try {
    # 檢查是否以管理員權限運行
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "此腳本需要管理員權限運行。請以管理員身份重新運行 PowerShell。"
    }

    # 建立 FwPolicy2 COM 對象
    $Rule = New-Object -ComObject FwPolicy2
    $AppLockerPolicy = $Rule.LocalPolicy.AppLockerPolicy

    # 檢查 AppLocker 服務是否運行
    $appLockerService = Get-Service -Name AppIDSvc -ErrorAction SilentlyContinue
    if ($appLockerService.Status -ne 'Running') {
        Write-Host "AppLocker 服務未運行。正在啟動服務..."
        Start-Service -Name AppIDSvc
    }

    # 建立新的 AppLocker 規則
    $PathRule = $AppLockerPolicy.NewAppLockerRule(3, "C:\Program Files\*", 0)
    $PathRule.Action = 1 # 允許
    $PathRule.Description = "允許所有用戶在 Program Files 目錄下運行程序"
    $PathRule.Name = "Allow Programs in Program Files"

    # 添加規則到政策
    $AppLockerPolicy.Add($PathRule)

    # 同樣添加一個規則для Program Files (x86)
    $PathRule86 = $AppLockerPolicy.NewAppLockerRule(3, "C:\Program Files (x86)\*", 0)
    $PathRule86.Action = 1 # 允許
    $PathRule86.Description = "允許所有用戶在 Program Files (x86) 目錄下運行程序"
    $PathRule86.Name = "Allow Programs in Program Files (x86)"
    $AppLockerPolicy.Add($PathRule86)

    # 應用新政策
    $Rule.LocalPolicy.AppLockerPolicy = $AppLockerPolicy

    # 保存變更
    $Rule.Save()

    Write-Host "AppLocker 政策已成功更新。"
}
catch {
    Write-Error "設置 AppLocker 政策時發生錯誤: $_"
    exit 1
}
finally {
    # 釋放 COM 對象
    if ($Rule) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rule) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
