@echo off
chcp 65001 >nul
echo ========================================
echo 파이썬 라이브러리 설치를 시작합니다...
echo ========================================
echo.

REM 파이썬이 설치되어 있는지 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] 파이썬이 설치되어 있지 않습니다.
    echo 파이썬을 먼저 설치해주세요: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo 파이썬 버전 확인:
python --version
echo.

REM pip 업그레이드
echo pip 업그레이드 중...
python -m pip install --upgrade pip
echo.

REM requirements.txt에서 라이브러리 설치
echo 필요한 라이브러리 설치 중...
python -m pip install -r requirements.txt
echo.

if errorlevel 1 (
    echo [오류] 라이브러리 설치 중 문제가 발생했습니다.
    pause
    exit /b 1
)

echo ========================================
echo 설치가 완료되었습니다!
echo ========================================
echo.
echo 이제 macro.py를 실행할 수 있습니다.
echo 실행 방법: python macro.py
echo.
pause

