@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem 콘솔 창 없이 GUI 실행 (pythonw 우선, 없으면 python)
set "LOCAL_PYTHONW=%LocalAppData%\Python\bin\pythonw.exe"
if exist "%LOCAL_PYTHONW%" (
    start "" "%LOCAL_PYTHONW%" "gui.py"
    exit /b
)

where py >nul 2>nul
if %errorlevel%==0 (
    start "" py -3w "gui.py"
    exit /b
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "gui.py"
) else (
    start "" python "gui.py"
)
