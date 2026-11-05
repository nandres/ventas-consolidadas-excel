@echo off
cd /d "%~dp0"
where py >nul 2>nul && (py consolidar.py) || (python consolidar.py)
echo.
pause
