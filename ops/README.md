# 발주봇 자동 실행 (Windows)

PC에 로그인하면 텔레그램 발주봇이 저절로 켜지고, 꺼지면 1분 뒤 다시 켜집니다.

## 설치
1. 이 폴더의 `install_bot_task.ps1`을 우클릭 → "PowerShell에서 실행".
2. "등록했고 지금 한 번 시작했습니다"가 나오면 끝입니다.

## 삭제
`uninstall_bot_task.ps1`을 같은 방법으로 실행합니다.

## 상태 확인
PowerShell에서:

    Get-ScheduledTask 발주봇 | Get-ScheduledTaskInfo

`LastTaskResult`가 `0`이거나 `267009`(실행 중)이면 정상입니다.

## 로그
`<데이터 폴더>\발주기록\로그\telegram_bot.log` (1 MB씩 5개까지 보관).
봇이 안 움직이면 이 파일의 마지막 줄을 먼저 보세요.

## 데이터 폴더 바꾸기
발주 기록·로그가 쌓이는 폴더입니다. 정하지 않으면 저장소 폴더의 바로 위 폴더를 씁니다.
저장소 폴더의 `.env`에 한 줄을 넣으면 바뀝니다 (예):

    RPA_DATA_DIR=C:\발주봇\데이터

바꾼 뒤에는 작업을 다시 시작하세요 (`Stop-ScheduledTask 발주봇` → `Start-ScheduledTask 발주봇`).

## 참고
- 봇이 이미 켜져 있으면 두 번째 봇은 "이미 실행 중"이라며 스스로 종료합니다.
- 같은 봇을 다른 PC에서도 켜면 텔레그램이 충돌(409)을 냅니다. 한 곳에서만 켜세요.
