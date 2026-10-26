# One-time (and re-runnable) setup of the order bot on the shop owner's PC. Started by 설치.bat.
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$envFile = Join-Path $repoRoot '.env'
$secretsFile = Join-Path $repoRoot '.streamlit\secrets.toml'
$taskName = '발주봇'
$dataDir = 'C:\choigoda-data'
$dataRepo = 'faulmann0435/Naver-RPA-data'
$dataBranch = 'dev'
# Shown on the first screen so a stale copy (old ZIP not overwritten) is easy to spot. Bump on every change.
$SetupEdition = '2026-10-26e'
$env:PYTHONUTF8 = '1'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

function Show-Step([int]$n, [string]$title) {
    Write-Host ''
    Write-Host "[$n/6] $title" -ForegroundColor Cyan
}
function Show-Ok([string]$m) { Write-Host "  OK  $m" -ForegroundColor Green }
function Show-Warn([string]$m) { Write-Host "  !!  $m" -ForegroundColor Yellow }
function Show-Fail([string]$m) { Write-Host "  XX  $m" -ForegroundColor Red }
function Stop-Setup([string]$m) {
    Show-Fail $m
    Write-Host ''
    Write-Host '설치를 멈췄습니다. 위 안내대로 고친 뒤 설치.bat을 다시 실행하세요. (이미 한 일은 다시 해도 괜찮습니다)' -ForegroundColor Red
    exit 1
}

function Ask-Value([string]$Label, [string]$Existing, [scriptblock]$Check, [string]$Hint) {
    # Empty input keeps the old value (never printed); otherwise the pasted value is cleaned and checked.
    while ($true) {
        if ($Existing) { $answer = Read-Host "$Label (이미 있음, Enter = 그대로)" }
        else { $answer = Read-Host $Label }
        $value = Clean-Value $answer
        if (-not $value -and $Existing) { return $Existing }
        if ($value -and (& $Check $value)) { return $value }
        Show-Warn $Hint
    }
}

function Ask-OptionalSecret([string]$Label, [string]$Existing) {
    # Masked input (nothing shows while typing). Enter skips; returns '' then.
    if ($Existing) { $Label = "$Label (이미 있음, Enter = 그대로)" }
    $secure = Read-Host $Label -AsSecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { return (Clean-Value ([System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr))) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

Write-Host '=== 발주봇 설치 ===' -ForegroundColor Cyan
Write-Host "설치 프로그램: ${SetupEdition}판"
Write-Host "설치 폴더: $repoRoot"

# ---- 1. Python -------------------------------------------------------------
Show-Step 1 '파이썬 확인'
$py = Get-PythonInfo
if (-not $py.Ok -and $py.Problem -match '너무 새것') {
    Show-Warn $py.Problem
    Write-Host "  발주봇용 파이썬 $PyRecommended 을 추가로 설치합니다. 지금 있는 파이썬은 그대로 둡니다." -ForegroundColor Yellow
    Write-Host '  (질문이 나오면 Y 또는 Enter를 누르세요. 1~2분 걸립니다)'
    if (Install-RecommendedPython) { $py = Get-PythonInfo }
}
if (-not $py.Ok) {
    Stop-Setup ("$($py.Problem)`n  python.org/downloads/release/python-3139/ 에서 파이썬 $PyRecommended 을 설치하세요." +
                " (맨 위 'Download Python 3.15' 큰 버튼이 아니라 3.13 버전입니다)" +
                "`n  설치 첫 화면에서 'Add python.exe to PATH'를 꼭 체크하고 Install Now를 누릅니다." +
                "`n  설치 후 이 창을 닫고 설치.bat을 다시 실행하세요.")
}
$python = $py.Path
Show-Ok "파이썬 $($py.Version) ($python)"

# ---- 2. Packages -----------------------------------------------------------
Show-Step 2 '필요한 부품 설치 (몇 분 걸릴 수 있습니다)'
Set-Location $repoRoot
$req = Join-Path $repoRoot 'requirements-bot.txt'
if (-not (Test-Path -LiteralPath $req)) { $req = Join-Path $repoRoot 'requirements.txt' }
& $python -m pip install --disable-pip-version-check -r $req
if ($LASTEXITCODE -ne 0) {
    Stop-Setup '부품 설치에 실패했습니다. 인터넷 연결을 확인하고 다시 실행하세요. 위쪽 빨간 글씨를 개발자에게 보여 주세요.'
}
Show-Ok '부품 설치 완료'

# ---- 3. IP ------------------------------------------------------------------
Show-Step 3 '이 PC의 인터넷 주소(IP) 확인 및 네이버 등록'
$ip = $null
try { $ip = (Invoke-RestMethod -Uri 'https://api.ipify.org' -TimeoutSec 15).ToString().Trim() } catch { }
if ($ip) {
    Write-Host ''
    Write-Host "  이 PC의 인터넷 주소(IP):  $ip" -ForegroundColor Yellow
    Write-Host '  이 숫자를 커머스API센터(apicenter.commerce.naver.com) > 내 애플리케이션 > API 호출 IP 에 추가해야 합니다.'
    Write-Host ''
} else {
    Show-Warn 'IP를 알아내지 못했습니다 (인터넷 확인). 구글에서 "내 IP"를 검색해 나오는 숫자를 등록하세요.'
}
Write-Host '  [지금 할 일] 인터넷 창에서 apicenter.commerce.naver.com 에 로그인합니다.' -ForegroundColor Cyan
Write-Host '   - 내 애플리케이션을 열어 둡니다. 다음 단계에서 "애플리케이션 ID"와 "시크릿"을 하나씩 복사해 붙여넣습니다.'
Write-Host '   - 같은 화면의 "API 호출 IP"에 위의 IP를 추가하고 저장합니다.'
Write-Host '  그 창을 닫지 말고, 준비가 되면 아래에서 Enter를 누르세요.'
Read-Host '  (Enter)' | Out-Null

# ---- 4. Settings -----------------------------------------------------------
Show-Step 4 '설정값 4개 입력 (붙여넣기 = 마우스 오른쪽 클릭, 그다음 Enter)'
$envMap = Read-EnvMap $envFile
$oldToken = ''
foreach ($line in (Read-TextLines $secretsFile)) {
    if ($line -match '^\s*token\s*=\s*"(.*)"\s*$') { $oldToken = $Matches[1] }
}
$notEmpty = { param($v) $v.Length -gt 0 }
$clientId = Ask-Value '1) 애플리케이션 ID  [커머스API센터에서 복사]' $envMap['NAVER_COMMERCE_CLIENT_ID'] $notEmpty '값이 비어 있습니다. 다시 붙여넣어 주세요.'
$clientSecret = Ask-Value '2) 시크릿  [커머스API센터에서 복사]' $envMap['NAVER_COMMERCE_CLIENT_SECRET'] $notEmpty '값이 비어 있습니다. 다시 붙여넣어 주세요.'
$botToken = Ask-Value '3) 텔레그램 봇 토큰  [BotFather가 알려 준 123456789:AA... 값]' $envMap['TELEGRAM_BOT_TOKEN'] { param($v) $v -match '^\d+:[\w-]+$' } '봇 토큰은 123456789:AAxxxx 모양입니다. 다시 붙여넣어 주세요.'
$ghToken = Ask-Value '4) GitHub 데이터 열람 토큰  [개발자에게 받은 github_pat_... 값]' $oldToken $notEmpty '값이 비어 있습니다. 다시 붙여넣어 주세요.'

Write-Host ''
Write-Host '[선택] 고객 문의 AI 답변 초안' -ForegroundColor Cyan
Write-Host '  Anthropic API 키를 넣으면 새 고객 문의마다 텔레그램으로 AI 초안과 버튼이 옵니다. 없으면 Enter로 건너뜁니다.'
Write-Host '  (텔레그램에서 [올리기]를 누르면 네이버에 바로 답변이 올라갑니다)'
$aiKey = Ask-OptionalSecret '5) Anthropic API 키  [sk-ant-... 값, 건너뛰려면 Enter]' $envMap['ANTHROPIC_API_KEY']

$chatIds = if ($envMap.Contains('TELEGRAM_ALLOWED_CHAT_IDS')) { $envMap['TELEGRAM_ALLOWED_CHAT_IDS'] } else { '' }
$envUpdates = [ordered]@{
    NAVER_COMMERCE_CLIENT_ID     = $clientId
    NAVER_COMMERCE_CLIENT_SECRET = $clientSecret
    TELEGRAM_BOT_TOKEN           = $botToken
    TELEGRAM_ALLOWED_CHAT_IDS    = $chatIds
    RPA_DATA_DIR                 = $dataDir
}
if ($aiKey) {
    $envUpdates['ANTHROPIC_API_KEY'] = $aiKey
    $envUpdates['CS_WATCH_ENABLED'] = '1'
}
# Real posting is on by default once there is an AI key, new or kept with Enter (the owner confirms every
# answer with [올리기]); a value already in .env is kept, so CS_ANSWER_ENABLED=0 survives a reinstall.
$hasAiKey = [bool]($aiKey -or $envMap['ANTHROPIC_API_KEY'])
if ($hasAiKey -and -not $envMap.Contains('CS_ANSWER_ENABLED')) { $envUpdates['CS_ANSWER_ENABLED'] = '1' }
Update-EnvFile $envFile $envUpdates
$tomlToken = $ghToken.Replace('\', '\\')
$toml = "[data_store]`r`nrepo = `"$dataRepo`"`r`nbranch = `"$dataBranch`"`r`ntoken = `"$tomlToken`"`r`n"
Write-Utf8NoBom $secretsFile $toml
if (-not (Test-Path -LiteralPath $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }
Show-Ok '설정을 저장했습니다 (.env, .streamlit\secrets.toml)'
Show-Ok "데이터 폴더: $dataDir"
if ($hasAiKey) {
    $answerOn = if ($envUpdates.Contains('CS_ANSWER_ENABLED')) { $envUpdates['CS_ANSWER_ENABLED'] } else { $envMap['CS_ANSWER_ENABLED'] }
    if ($answerOn -eq '1') { Show-Ok '고객 문의 알림을 켰습니다 ([올리기] = 네이버에 실제 등록)' }
    else { Show-Ok '고객 문의 알림을 켰습니다 (답변 올리기는 시험 모드)' }
}

# ---- 5. Naver + Telegram --------------------------------------------------------------
Show-Step 5 '네이버·텔레그램 연결 확인'
$naverText = (& $python (Join-Path $PSScriptRoot 'check_naver.py') 2>&1 | Out-String).Trim()
$naverExit = $LASTEXITCODE
$naverOut = ($naverText -split '\r?\n' | Select-Object -Last 1)
$naverOk = ($naverExit -eq 0 -and $naverOut -eq 'OK')
if ($naverOk) {
    Show-Ok '네이버 연결 OK'
} elseif ($naverText -match 'IP_NOT_ALLOWED') {
    Show-Warn '네이버가 이 PC의 IP를 아직 허용하지 않았습니다.'
    Write-Host '  위의 IP를 커머스API센터에 등록한 뒤(반영에 몇 분), 설치.bat을 다시 실행하면 확인됩니다.'
    Write-Host '  지금은 설치를 계속 진행합니다.'
} else {
    Show-Warn "네이버 연결 실패: $naverOut"
    Write-Host '  애플리케이션 ID/시크릿이 맞는지 확인하세요. 값을 고치려면 설치.bat을 다시 실행합니다. 지금은 설치를 계속 진행합니다.'
}

$botName = $null
try {
    $me = Invoke-RestMethod -Uri "https://api.telegram.org/bot$botToken/getMe" -TimeoutSec 20
    if ($me.ok) { $botName = $me.result.username }
} catch {
    $status = $null
    if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
    if ($status -eq 401 -or $status -eq 404) { Show-Fail '텔레그램 봇 토큰이 틀렸습니다. 설치.bat을 다시 실행해 토큰을 새로 붙여넣어 주세요.' }
    else { Show-Fail '텔레그램에 연결하지 못했습니다. 인터넷 연결을 확인하고 설치.bat을 다시 실행하세요.' }
}
if ($botName) { Show-Ok "텔레그램 봇 확인: @$botName" }

# ---- 6. Auto start ---------------------------------------------------------
Show-Step 6 '컴퓨터를 켜면 봇이 저절로 켜지게 등록'
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue  # so the new settings are read
}
& (Join-Path $PSScriptRoot 'install_bot_task.ps1')
$taskOk = ($LASTEXITCODE -eq 0 -or $null -eq $LASTEXITCODE) -and [bool](Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)
if (-not $taskOk) {
    Show-Fail '자동 실행 등록에 실패했습니다. 설치.bat을 마우스 오른쪽 클릭 > "관리자 권한으로 실행"으로 다시 해 보세요.'
}

Write-Host ''
Write-Host '절전 모드가 켜지면 봇도 같이 멈춥니다. 전원 연결 시 "절전 모드: 안 함"으로 바꿀까요?' -ForegroundColor Yellow
$sleepAnswer = Read-Host '바꾸려면 Y, 건너뛰려면 Enter'
if ($sleepAnswer -match '^[yYㅛ]') {
    & powercfg /change standby-timeout-ac 0
    & powercfg /change hibernate-timeout-ac 0
    if ($LASTEXITCODE -eq 0) { Show-Ok '절전 모드를 "안 함"으로 바꿨습니다.' }
    else { Show-Warn '바꾸지 못했습니다. Windows 설정 > 시스템 > 전원에서 직접 "절전 모드: 안 함"으로 바꿔 주세요.' }
} else {
    Write-Host '  건너뜁니다. Windows 설정 > 시스템 > 전원에서 "절전 모드: 안 함"으로 직접 바꿔 주세요.'
}

# ---- Finish ----------------------------------------------------------------
Write-Host ''
Write-Host '=== 설치가 끝났습니다 ===' -ForegroundColor Green
if (-not $naverOk) { Write-Host '(네이버 연결은 아직 확인되지 않았습니다. 위 안내를 해결한 뒤 설치.bat을 다시 실행하세요.)' -ForegroundColor Yellow }
Write-Host ''
Write-Host '다음 순서:' -ForegroundColor Cyan
if ($botName) { Write-Host "  1. 휴대폰 텔레그램에서 @$botName 을 찾아 대화방을 엽니다." }
else { Write-Host '  1. 휴대폰 텔레그램에서 내 봇을 찾아 대화방을 엽니다.' }
Write-Host '  2. [시작] 버튼을 누릅니다. 봇이 "이 대화방 번호는 ○○ 입니다"라고 답합니다.'
Write-Host '  3. 그 번호를 확인한 뒤, 이 폴더의 "대화방추가.bat"을 더블클릭하고 번호를 입력합니다.'
