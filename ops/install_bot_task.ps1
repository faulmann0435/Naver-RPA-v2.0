# Registers the Telegram order bot as a Task Scheduler task that starts at logon.
$ErrorActionPreference = 'Stop'
$taskName = '발주봇'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host 'python을 찾을 수 없습니다. Python을 설치하고 PATH에 추가한 뒤 다시 실행하세요.' -ForegroundColor Red
    exit 1
}
$pythonw = Join-Path (Split-Path $python.Source) 'pythonw.exe'
if (-not (Test-Path $pythonw)) {
    Write-Host "pythonw.exe를 찾을 수 없습니다: $pythonw" -ForegroundColor Red
    exit 1
}

$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$action = New-ScheduledTaskAction -Execute $pythonw -Argument '-m worker.telegram_bot' -WorkingDirectory $repoRoot
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$settings = New-ScheduledTaskSettingsSet `
    -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null
Start-ScheduledTask -TaskName $taskName

Write-Host "등록했고 지금 한 번 시작했습니다: $taskName" -ForegroundColor Green
Write-Host "로그 파일: <데이터 폴더>\발주기록\로그\telegram_bot.log"
Write-Host '(데이터 폴더는 .env의 RPA_DATA_DIR, 없으면 저장소 폴더의 바로 위 폴더)'
