# Stops the Telegram order bot and removes its Task Scheduler task.
$ErrorActionPreference = 'Stop'
$taskName = '발주봇'

if (-not (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)) {
    Write-Host "등록된 작업이 없습니다: $taskName"
    exit 0
}
Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
Write-Host "작업을 중지하고 삭제했습니다: $taskName" -ForegroundColor Green
