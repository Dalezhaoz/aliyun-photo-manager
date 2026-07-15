@echo off
setlocal

cd /d "%~dp0"

echo [1/3] Preparing ARM64 Windows virtual environment...
if not exist ".venv\Scripts\python.exe" (
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3.12 -m venv .venv
        if errorlevel 1 py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Failed to create .venv. Please install Python 3.12 and try again.
    pause
    exit /b 1
)

echo [2/3] Installing dependencies...
".venv\Scripts\python.exe" -m pip install -U pip
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed

echo [3/3] Building ARM64 Windows self-use app...
set "ARM64_OUTPUT_ROOT=%LOCALAPPDATA%\aliyun_photo_manager_arm64"
".venv\Scripts\python.exe" build_windows_qt_app.py --channel stable --output-root "%ARM64_OUTPUT_ROOT%"
if errorlevel 1 goto failed

echo.
echo ARM64 app generated under:
echo %ARM64_OUTPUT_ROOT%\dist\aliyun_photo_manager_qt
echo.
start "" "%ARM64_OUTPUT_ROOT%\dist\aliyun_photo_manager_qt"
pause
exit /b 0

:failed
echo.
echo ARM64 release build failed. Please check the error messages above.
pause
exit /b 1
