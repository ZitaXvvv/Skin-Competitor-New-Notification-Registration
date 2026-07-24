<#
.SYNOPSIS
    注册两个 Windows 计划任务：
      1. CI_NewSKU_Monthly_Scrape  — 每月1号凌晨02:00，全量抓取过去40天新品
      2. CI_NewSKU_Daily_Fill      — 每天凌晨03:00，每次补全10个不完整产品卡片

.NOTES
    以管理员身份在项目根目录执行：.\scripts\setup_daily_fill_task.ps1
    Python 路径自动检测（取当前 shell 的 python 可执行路径）。
#>

param(
    [string]$PythonExe   = (Get-Command python -ErrorAction Stop).Source,
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot),
    [int]$DailyLimit     = 10,    # 每天补全的 SKU 上限
    [int]$MonthlyDays    = 40     # 月度抓取的回溯天数
)

$MainScript   = Join-Path $ProjectRoot "src\main.py"
$FillScript   = Join-Path $ProjectRoot "src\module6_daily_fill.py"
$LogDir       = Join-Path $ProjectRoot "log"

# ── 检查文件存在 ────────────────────────────────────────────────────────
foreach ($f in @($MainScript, $FillScript)) {
    if (-not (Test-Path $f)) {
        Write-Error "找不到脚本: $f"; exit 1
    }
}
New-Item -ItemType Directory -Force $LogDir | Out-Null

$TaskUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# ════════════════════════════════════════════════════════════════════════
# 任务1：月度全量抓取（原有，更新 --days 为40）
# ════════════════════════════════════════════════════════════════════════
$MonthlyName    = "CI_NewSKU_Monthly_Scrape"
$MonthlyTrigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "02:00"
$MonthlyAction  = New-ScheduledTaskAction `
    -Execute    $PythonExe `
    -Argument   "`"$MainScript`" --days $MonthlyDays --unattended" `
    -WorkingDirectory $ProjectRoot
$MonthlySettings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew

if (Get-ScheduledTask -TaskName $MonthlyName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $MonthlyName -Confirm:$false
    Write-Host "  [更新] $MonthlyName"
} else {
    Write-Host "  [新建] $MonthlyName"
}
Register-ScheduledTask `
    -TaskName   $MonthlyName `
    -Trigger    $MonthlyTrigger `
    -Action     $MonthlyAction `
    -Settings   $MonthlySettings `
    -RunLevel   Highest `
    -User       $TaskUser `
    -Force | Out-Null

Write-Host "  ✅ $MonthlyName — 每月1号 02:00，回溯 $MonthlyDays 天"

# ════════════════════════════════════════════════════════════════════════
# 任务2：每日增量补全（新增）
# ════════════════════════════════════════════════════════════════════════
$DailyName     = "CI_NewSKU_Daily_Fill"
$DailyTrigger  = New-ScheduledTaskTrigger -Daily -At "03:00"
$DailyAction   = New-ScheduledTaskAction `
    -Execute    $PythonExe `
    -Argument   "`"$FillScript`" --limit $DailyLimit" `
    -WorkingDirectory $ProjectRoot
$DailySettings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew

if (Get-ScheduledTask -TaskName $DailyName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $DailyName -Confirm:$false
    Write-Host "  [更新] $DailyName"
} else {
    Write-Host "  [新建] $DailyName"
}
Register-ScheduledTask `
    -TaskName   $DailyName `
    -Trigger    $DailyTrigger `
    -Action     $DailyAction `
    -Settings   $DailySettings `
    -RunLevel   Highest `
    -User       $TaskUser `
    -Force | Out-Null

Write-Host "  ✅ $DailyName — 每天 03:00，每次补全最多 $DailyLimit 个 SKU"

# ════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "=== 注册完成 ===" -ForegroundColor Green
Write-Host "任务计划程序中可查看：taskschd.msc"
Write-Host ""
Write-Host "手动立即测试连通性："
Write-Host "  python src\module6_daily_fill.py --check"
Write-Host "手动立即运行一次补全（dry-run，不写入）："
Write-Host "  python src\module6_daily_fill.py --dry-run"
Write-Host "手动立即运行一次补全（实际写入）："
Write-Host "  python src\module6_daily_fill.py"
Write-Host ""
Write-Host "日志位置: $LogDir\daily_fill_YYYYMMDD.log"
