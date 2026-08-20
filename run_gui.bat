@echo off
chcp 65001 > nul
cd /d "%~dp0"

REM 엑셀 매크로 봇 — GUI 실행
REM 처음 한 번만: pip install -r requirements.txt

start "" pythonw main.py

REM pythonw 가 없거나 창이 안 뜨면 아래 줄을 대신 쓰세요 (오류 메시지 확인 가능).
REM python main.py
REM pause
