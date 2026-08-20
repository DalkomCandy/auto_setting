@echo off
chcp 65001 > nul
cd /d "%~dp0"

REM ─────────────────────────────────────────────
REM  진단 도구 (CLI) — 본 작업은 run_gui.bat 을 쓰세요.
REM  이 파일은 값을 직접 채워 넣지 않고 매번 인자로 받습니다.
REM  그래야 나중에 git pull 로 이 파일이 갱신돼도 충돌이 나지 않습니다.
REM
REM  사용법:
REM    run.bat                                          → 매크로 이름 목록만 출력
REM    run.bat "D:\all폴더" "PERSONAL.XLSB!수익_개요"     → 목록 출력 + 첫 파일로 대화상자 구조 확인
REM ─────────────────────────────────────────────

echo.
echo   [1] 사용 가능한 매크로 이름 보기
echo.
python excel_macro_bot.py --list-macros

if "%~1"=="" goto :end
if "%~2"=="" (
    echo.
    echo   두 번째 인자로 매크로 이름도 넘겨야 대화상자 구조를 확인합니다.
    echo   예: run.bat "%~1" "PERSONAL.XLSB!수익_개요"
    goto :end
)

echo.
echo   [2] 매크로가 띄우는 창의 구조 보기 (저장하지 않음)
echo       대상 폴더 : %~1
echo       매크로    : %~2
echo       실행 전에 Excel 을 모두 닫아주세요.
echo.
pause
python excel_macro_bot.py --root "%~1" --macro "%~2" --probe --limit 1

:end
echo.
pause
