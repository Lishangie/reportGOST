@echo off
rem Rebuilds all examples and checks there are no remarks.
chcp 65001 >nul
cd /d "%~dp0"
set "PY=python"
where python >nul 2>nul || set "PY=py"
%PY% check.py
echo.
pause
