@echo off
rem Layout debugging: drafts and their PDFs are kept in _otladka folder next to report.md
chcp 65001 >nul
set "SRC=%~1"
if "%SRC%"=="" set "SRC=%~dp0report.md"
call "%~dp0Собрать.bat" "%SRC%" --keep-temp
