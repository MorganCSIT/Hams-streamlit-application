@echo off
cd /d "%~dp0"

if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" (
    "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" desktop_launcher.py
) else if exist "%LOCALAPPDATA%\Python\bin\python.exe" (
    "%LOCALAPPDATA%\Python\bin\python.exe" desktop_launcher.py
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 desktop_launcher.py
    ) else (
        python desktop_launcher.py
    )
)
