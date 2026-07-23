<#
.SYNOPSIS
  在 Windows 服务器上注册"每月自动跑一次全量抓取"的计划任务。
  只需在服务器上以管理员身份运行一次即可（不是每次部署都要跑）。

.用法
  以管理员身份打开 PowerShell，cd 到本仓库根目录，执行：
    .\scripts\setup_monthly_task.ps1
  可选参数：
    -Day <1-28>      每月第几天运行，默认 1 号
    -Time "HH:mm"    几点运行，默认 02:00（凌晨，避免占用高峰网络/CPU）
    -Days <N>        抓取回溯天数（传给 main.py --days），默认 31（略大于一个月，避免漏单）

.说明
  任务本体就是: python src\main.py --days <Days>
  跑完后数据写入 CI_List_Ada.xlsx，Streamlit 看板读的是同一份文件，下次有人刷新页面
  （或缓存 30 分钟过期后）就能看到最新数据，不需要重启 streamlit 进程。
#>
param(
    [ValidateRange(1,28)][int]$Day = 1,
    [string]$Time = "02:00",
    [int]$Days = 31,
    [string]$TaskName = "CI_NewSKU_Monthly_Scrape"
)

$RepoRoot = Split-Path -Parent $PSScriptRoot
$PythonExe = (Get-Command python).Source
$MainScript = Join-Path $RepoRoot "src\main.py"

if (-not (Test-Path $MainScript)) {
    Write-Error "找不到 $MainScript，请确认在正确的仓库目录下运行本脚本"
    exit 1
}

$Action  = New-ScheduledTaskAction -Execute $PythonExe `
    -Argument "`"$MainScript`" --days $Days" `
    -WorkingDirectory $RepoRoot

$Trigger = New-ScheduledTaskTrigger -Monthly -At $Time -DaysOfMonth $Day

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopOnIdleEnd `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `
    -RestartCount 1 -RestartInterval (New-TimeSpan -Minutes 10)

Register-ScheduledTask -TaskName $TaskName `
    -Action $Action -Trigger $Trigger -Settings $Settings `
    -Description "CI New SKU Cross Brand：每月自动跑一次抓取流程（main.py），补全竞品新品数据" `
    -RunLevel Highest -Force

Write-Output "✅ 已注册计划任务 [$TaskName]：每月 $Day 号 $Time 自动运行 `"python src\main.py --days $Days`""
Write-Output "   可以在 '任务计划程序' (taskschd.msc) 里查看/手动测试运行该任务"
