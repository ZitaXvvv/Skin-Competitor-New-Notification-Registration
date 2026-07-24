<#
.SYNOPSIS
  交互式 BEBD 登录刷新工具。
  在任意可以操作鼠标/键盘的工作时段运行此脚本，完成一次登录后
  Cookie 自动保存，当晚凌晨的计划任务可直接复用，无需人工值守。

.用法
  在项目根目录执行（无需管理员身份）：
    .\scripts\refresh_bebd_login.ps1

  建议在每月计划任务运行前（月底最后一天的工作时间）执行一次，
  或在每日日志出现 "BEBD: ❌" 时立即执行。
#>

param(
    [string]$PythonExe = (Get-Command python -ErrorAction Stop).Source,
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)

$LoginScript = Join-Path $ProjectRoot "src\refresh_bebd_login.py"

if (-not (Test-Path $LoginScript)) {
    Write-Error "找不到登录脚本: $LoginScript"; exit 1
}

Write-Host ""
Write-Host "BEBD 登录刷新工具" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "将打开一个可见的浏览器窗口，请："
Write-Host "  1. 在弹出的浏览器里登录美丽修行大数据 (bebd.bevol.com)"
Write-Host "  2. 登录成功后回到此终端，按 Enter 保存 Cookie"
Write-Host ""
Write-Host "保存后，今晚凌晨的计划任务将自动复用此 Cookie。" -ForegroundColor Green
Write-Host ""

& $PythonExe $LoginScript

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "✅ Cookie 已保存，凌晨计划任务无需人工干预。" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "❌ 登录保存失败，请检查网络或重试。" -ForegroundColor Red
}
