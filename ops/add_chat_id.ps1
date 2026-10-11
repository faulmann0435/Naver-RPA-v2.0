# Adds a Telegram chat number to TELEGRAM_ALLOWED_CHAT_IDS in .env and restarts the bot. Started by 대화방추가.bat.
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$envFile = Join-Path $repoRoot '.env'
$taskName = '발주봇'
$key = 'TELEGRAM_ALLOWED_CHAT_IDS'

Write-Host '=== 텔레그램 대화방 추가 ===' -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $envFile)) {
    Write-Host '설정 파일(.env)이 없습니다. 먼저 설치.bat을 실행하세요.' -ForegroundColor Red
    exit 1
}

Write-Host '봇이 알려 준 "이 대화방 번호는 ○○ 입니다"의 숫자를 입력합니다.'
$id = ''
while ($true) {
    $id = Clean-Value (Read-Host '대화방 번호')
    if ($id -match '^-?\d+$') { break }
    Write-Host '숫자만 입력하세요 (앞에 -가 붙은 번호는 -까지 같이). 다시 입력해 주세요.' -ForegroundColor Yellow
}

$map = Read-EnvMap $envFile
$current = @()
if ($map.Contains($key) -and $map[$key]) {
    $current = @($map[$key] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}
if ($current -contains $id) {
    Write-Host "이미 등록된 번호입니다: $id" -ForegroundColor Yellow
} else {
    $current = @($current) + $id
    Update-EnvFile $envFile ([ordered]@{ $key = ($current -join ',') })
    Write-Host "추가했습니다: $id" -ForegroundColor Green
}

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-ScheduledTask -TaskName $taskName
    Write-Host '봇을 다시 시작했습니다.' -ForegroundColor Green
} else {
    Write-Host '등록된 봇 작업이 없습니다. 설치.bat을 먼저 실행하세요.' -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host '이제 휴대폰 텔레그램으로 돌아가 봇에게 아무 글자나 보내 보세요.' -ForegroundColor Cyan
Write-Host '화면 아래에 [📦 발주 처리] 버튼이 나타나면 성공입니다. (1분쯤 걸릴 수 있습니다)'
