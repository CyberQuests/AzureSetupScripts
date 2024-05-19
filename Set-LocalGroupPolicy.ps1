# Set-LocalGroupPolicy.ps1

# 允許用戶在 Program Files 目錄下安裝應用程式
$Rule = New-Object -ComObject FwPolicy2
$AppLockerPolicy = $Rule.LocalPolicy.AppLockerPolicy

# 新增一個規則允許所有使用者在 Program Files 目錄下運行應用程式
$PathRule = $AppLockerPolicy.NewAppLockerRule(3, "C:\Program Files\*", 0)
$PathRule.Action = 1
$PathRule.Description = "Allow all users to run programs in Program Files"

# 添加規則
$AppLockerPolicy.Add($PathRule)
$Rule.LocalPolicy.AppLockerPolicy = $AppLockerPolicy

# 保存變更
$Rule.Save()
