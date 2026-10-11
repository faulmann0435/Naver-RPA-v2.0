# 발주봇 자동 실행 (Windows)

PC에 로그인하면 텔레그램 발주봇이 저절로 켜지고, 꺼지면 1분 뒤 다시 켜집니다.

## 처음 설치 (사장님 PC)
1. 저장소 ZIP을 `C:\`에 풀면 `C:\Naver-RPA-v2.0-order-bot-test` 가 생깁니다.
2. 그 폴더의 `설치.bat`을 더블클릭 -> `ops\setup_owner_pc.ps1`이 6단계로 진행합니다
   (파이썬 확인 / 부품 설치 / IP 확인·네이버 등록 / 설정값 4개 입력 / 연결 확인 / 자동 실행 등록).
   다시 실행해도 안전합니다. 이미 있는 값은 Enter로 그대로 둡니다.
3. 텔레그램에서 봇에게 [시작] -> 봇이 알려 준 대화방 번호를 `대화방추가.bat`(`ops\add_chat_id.ps1`)에 입력.

값의 출처: 애플리케이션 ID·시크릿은 커머스API센터, 봇 토큰은 @BotFather, GitHub 토큰만 개발자.
봇에는 `requirements-bot.txt`(streamlit/anthropic 제외)만 설치합니다.
`ops\_common.ps1`은 공용 함수(.env 읽기/쓰기, 파이썬 찾기), `ops\check_naver.py`는 네이버 토큰 시험입니다.

## 작업 스케줄러만 따로 등록
`install_bot_task.ps1`을 우클릭 -> "PowerShell에서 실행". 파이썬 경로는 `python` 자신에게 물어 찾습니다
(Microsoft Store용 가짜 python은 건너뜁니다).

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
