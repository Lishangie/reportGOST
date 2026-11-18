@echo off
rem Build report: double-click (builds report.md next to this file),
rem drag a report.md onto this icon, or call:  Sobrat.bat path\to\report.md
chcp 65001 >nul
set "SRC=%~1"
if "%SRC%"=="" set "SRC=%~dp0report.md"

set "PY=python"
where python >nul 2>nul || set "PY=py"
%PY% --version >nul 2>nul || (
    echo Python not found. Install Python 3.10+ from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" in the installer.
    echo.
    pause
    exit /b 1
)

%PY% -c "import docx, PIL, lxml" >nul 2>nul || (
    echo Installing python-docx, pillow, lxml ...
    %PY% -m pip install --quiet python-docx pillow lxml
)
%PY% -c "import pymupdf" >nul 2>nul || (
    echo Installing pymupdf ...
    %PY% -m pip install --quiet pymupdf
)
where pandoc >nul 2>nul || %PY% -c "import pypandoc; pypandoc.get_pandoc_path()" >nul 2>nul || (
    echo Installing pandoc ^(pypandoc_binary^) for formulas ...
    %PY% -m pip install --quiet pypandoc_binary
)

%PY% "%~dp0build.py" "%SRC%" %2 %3 %4
echo.
if not "%NOPAUSE%"=="1" pause
