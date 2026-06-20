@echo off
setlocal

set "PROJECT_DIR=%~dp0"
cd /d "%PROJECT_DIR%"

title Citation Notifier

echo ============================================
echo   Citation Notifier - Launcher
echo ============================================
echo.

if /i "%~1"=="--check" goto CHECK_ONLY

call :CHECK_ENV
if errorlevel 1 goto FAIL

echo [4/4] Starting web server...
echo URL: http://127.0.0.1:5000
echo Close this window to stop the server.
echo.

python main.py web
echo.
echo Server exited.
pause
exit /b 0

:CHECK_ONLY
call :CHECK_ENV
exit /b %errorlevel%

:CHECK_ENV
echo [1/4] Checking project files...
if not exist "main.py" (
    echo ERROR: main.py was not found.
    echo Current directory: %CD%
    echo Put this launcher in the project root folder.
    exit /b 1
)
if not exist "requirements.txt" (
    echo ERROR: requirements.txt was not found.
    echo Current directory: %CD%
    exit /b 1
)
echo Project files OK.
echo.

echo [2/4] Checking Python...
python --version >nul 2>nul
if errorlevel 1 (
    echo ERROR: python command is not available.
    echo Install Python 3.10 or newer and add it to PATH.
    exit /b 1
)
python --version
echo.

echo [3/4] Checking dependencies...
python -c "import flask, yaml, pandas, docx, pdfplumber, openpyxl, xlrd, jinja2" >nul 2>nul
if errorlevel 1 (
    echo Dependencies are missing. Installing requirements.txt...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: dependency installation failed.
        echo Try manually: python -m pip install -r requirements.txt
        exit /b 1
    )
    echo Dependencies installed.
) else (
    echo Dependencies OK.
)
echo.

echo [4/4] Checking runtime folders...
if not exist "data" mkdir "data"
if not exist "data\input" mkdir "data\input"
echo Runtime folders OK.
echo.
exit /b 0

:FAIL
echo.
echo Launch failed. See messages above.
pause
exit /b 1
