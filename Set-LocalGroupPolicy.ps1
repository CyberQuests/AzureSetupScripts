# Set-LocalGroupPolicy.ps1
# 此腳本用於修改 AppLocker 政策，允許用戶在 Program Files 目錄下安裝和運行應用程式

# 設置嚴格的錯誤處理，遇到錯誤時停止執行
$ErrorActionPreference = "Stop"

try {
    # 檢查是否以管理員權限運行
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "此腳本需要管理員權限運行。請以管理員身份重新運行 PowerShell。"
    }

    # 檢查 AppLocker 服務是否運行
    $appLockerService = Get-Service -Name AppIDSvc -ErrorAction SilentlyContinue
    if ($appLockerService.Status -ne 'Running') {
        Write-Host "AppLocker 服務未運行。正在啟動服務..."
        Start-Service -Name AppIDSvc
    }

    # 檢查是否有 Set-AppLockerPolicy cmdlet
    if (-not (Get-Command Set-AppLockerPolicy -ErrorAction SilentlyContinue)) {
        throw "找不到 Set-AppLockerPolicy cmdlet。請確保 AppLocker 模組已安裝。"
    }

    # 定義 AppLocker 政策的 XML 內容
    $appLockerPolicyXml = @"
<AppLockerPolicy Version="1">
  <RuleCollection Type="Exe" EnforcementMode="Enabled">
    <FilePathRule Id="AllowProgramsInProgramFiles" Name="Allow Programs in Program Files" Description="允許所有用戶在 Program Files 目錄下運行程序" UserOrGroupSid="S-1-1-0" Action="Allow" Priority="1">
      <Conditions>
        <FilePathCondition Path="C:\Program Files\*" />
      </Conditions>
    </FilePathRule>
    <FilePathRule Id="AllowProgramsInProgramFilesx86" Name="Allow Programs in Program Files (x86)" Description="允許所有用戶在 Program Files (x86) 目錄下運行程序" UserOrGroupSid="S-1-1-0" Action="Allow" Priority="2">
      <Conditions>
        <FilePathCondition Path="C:\Program Files (x86)\*" />
      </Conditions>
    </FilePathRule>
  </RuleCollection>
</AppLockerPolicy>
"@

    # 將 AppLocker 政策寫入臨時文件
    $tempPolicyPath = "$env:TEMP\AppLockerPolicy.xml"
    Write-Host "正在建立 AppLocker 政策文件..."
    $appLockerPolicyXml | Out-File -FilePath $tempPolicyPath -Encoding UTF8

    # 應用 AppLocker 政策，使用 Merge 模式以保留現有政策並添加新規則
    Write-Host "正在應用 AppLocker 政策..."
    Set-AppLockerPolicy -XmlPolicy $tempPolicyPath -Merge

    # 刪除臨時政策文件
    Remove-Item -Path $tempPolicyPath -Force

    Write-Host "AppLocker 政策已成功更新。"
}
catch {
    Write-Error "設置 AppLocker 政策時發生錯誤: $_"
    exit 1
}
finally {
    # 釋放 COM 對象（如果有的話）
    if ($Rule) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rule) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
