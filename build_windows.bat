@echo off
REM Build Tansu.exe for Windows

echo ===================================
echo Building Tansu for Windows
echo ===================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python 3.8+
    pause
    exit /b 1
)

REM Get script directory
cd /d "%~dp0"

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller pystray Pillow

REM Clean previous build
echo Cleaning previous build...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM Build the app
echo Building application...
pyinstaller --noconfirm Tansu.spec

REM Check if build succeeded
if exist "dist\Tansu\Tansu.exe" (
    echo.
    echo ===================================
    echo Build successful!
    echo ===================================
    echo.
    echo App location: dist\Tansu\Tansu.exe
    echo.
    echo To run: double-click dist\Tansu\Tansu.exe
    echo.
    echo For Word ribbon integration:
    echo   1. Start Tansu first (it runs the API server)
    echo   2. Follow word_addin\INSTALL_WORD_ADDIN.txt
    echo.
) else (
    echo.
    echo ===================================
    echo Build failed!
    echo ===================================
    echo Check the output above for errors.
    pause
    exit /b 1
)

pause
