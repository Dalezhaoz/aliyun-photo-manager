@echo off
setlocal

cd /d "%~dp0"

echo [1/3] Preparing Windows virtual environment...
if not exist ".venv\Scripts\python.exe" (
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3.12-64 -m venv .venv
        if errorlevel 1 py -3-64 -m venv .venv
    ) else (
        python -m venv .venv
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Failed to create .venv. Please install Python 3.12 and try again.
    pause
    exit /b 1
)

echo Checking Python architecture...
".venv\Scripts\python.exe" -c "import platform, sys; arch=platform.machine().upper(); print('Python:', sys.version.split()[0], arch); raise SystemExit(0 if arch in ('AMD64','X86_64') else 1)"
if errorlevel 1 (
    echo.
    echo This Python is not AMD64/x64. A package built here may not run on normal Intel/AMD Windows PCs.
    echo Please build with x64 Python on an x64 Windows machine, or install x64 Python and recreate .venv.
    pause
    exit /b 1
)

echo [2/3] Installing dependencies...
".venv\Scripts\python.exe" -m pip install -U pip
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed

echo [3/3] Building test app directory...
".venv\Scripts\python.exe" build_windows_qt_app.py --channel stable
if errorlevel 1 goto failed

echo.
echo Test build generated:
echo %cd%\dist_qt\aliyun_photo_manager_qt
echo.
start "" "%cd%\dist_qt\aliyun_photo_manager_qt"
pause
exit /b 0

:failed
echo.
echo Test build failed. Please check the error messages above.
pause
exit /b 1
