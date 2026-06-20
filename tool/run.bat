@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."
echo ============================================
echo  수량산출서 출력 자동화 도구
echo ============================================
set /p TARGET="처리할 폴더 또는 파일 경로를 입력하세요: "
python "tool\toolruntime.py" "%TARGET%"
echo.
pause
