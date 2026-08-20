@echo off
chcp 65001 > nul
cd /d "%~dp0"

REM ─────────────────────────────────────────────
REM  진단 도구 (CLI) — 본 작업은 run_gui.bat 을 쓰세요.
REM ─────────────────────────────────────────────
set ROOT=D:\all폴더
set MACRO=PERSONAL.XLSB!수익_개요

echo.
echo   [1] 사용 가능한 매크로 이름 보기
echo.
python excel_macro_bot.py --list-macros

echo.
echo   [2] 매크로가 띄우는 창의 구조 보기 (저장하지 않음)
echo       실행 전에 Excel 을 모두 닫아주세요.
echo.
pause
python excel_macro_bot.py --root "%ROOT%" --macro "%MACRO%" --probe --limit 1

echo.
pause
