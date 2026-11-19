@echo off
rem Updates build.py and docs in this folder from the shared kit folder
rem (the nearest folder with build.py and primery/ found above this one).
chcp 65001 >nul
set "PY=python"
where python >nul 2>nul || set "PY=py"
%PY% "%~dp0build.py" --update
echo.
pause