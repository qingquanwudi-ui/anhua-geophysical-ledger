@echo off
setlocal EnableExtensions

rem Resolve paths from this script location so the project can be moved.
set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"
set "APP_PATH=%PROJECT_ROOT%\work\ledger_web_app.py"
set "LOG_DIR=%PROJECT_ROOT%\outputs\ledger_system"

if not exist "%APP_PATH%" (
    echo Cannot find Web app: "%APP_PATH%"
    exit /b 1
)

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

rem Optional override: set LEDGER_PYTHON_EXE before running this script if needed.
if defined LEDGER_PYTHON_EXE (
    set "PYTHON_EXE=%LEDGER_PYTHON_EXE%"
    set "PYTHON_ARGS="
)

rem Prefer the bundled Codex runtime when this project is run from Codex Desktop.
if not defined PYTHON_EXE (
    if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
        set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
        set "PYTHON_ARGS="
    )
)

if not defined PYTHON_EXE (
    where python.exe >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_EXE=python.exe"
        set "PYTHON_ARGS="
    )
)

if not defined PYTHON_EXE (
    where py.exe >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_EXE=py.exe"
        set "PYTHON_ARGS=-3"
    )
)

if not defined PYTHON_EXE (
    echo Python was not found. Install Python or set LEDGER_PYTHON_EXE before running this script.
    exit /b 1
)

if defined LEDGER_DRY_RUN (
    echo PROJECT_ROOT=%PROJECT_ROOT%
    echo APP_PATH=%APP_PATH%
    echo LOG_DIR=%LOG_DIR%
    echo PYTHON_EXE=%PYTHON_EXE%
    echo PYTHON_ARGS=%PYTHON_ARGS%
    exit /b 0
)

cd /d "%PROJECT_ROOT%"
"%PYTHON_EXE%" %PYTHON_ARGS% "%APP_PATH%" >> "%LOG_DIR%\server_stdout.log" 2>> "%LOG_DIR%\server_stderr.log"