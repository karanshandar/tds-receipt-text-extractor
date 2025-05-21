@echo off
echo ===============================================================
echo               PDF Extractor - Setup
echo ===============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Setting up PDF Extractor...

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        echo Please make sure you have the venv module installed:
        echo pip install virtualenv
        pause
        exit /b 1
    )
    echo Virtual environment created successfully.
) else (
    echo Virtual environment already exists.
)

REM Activate virtual environment and install packages
echo Activating virtual environment and installing required packages...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Failed to activate virtual environment.
    pause
    exit /b 1
)

echo Installing required packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install required packages.
    pause
    exit /b 1
)

REM Create input directory if it doesn't exist
if not exist "input" (
    echo Creating input directory...
    mkdir input
)

echo.
echo Setup completed successfully!
echo.
echo To use PDF Extractor:
echo 1. Place your PDF files in the 'input' folder
echo 2. Run 'run_extractor.bat' to process them
echo.
echo Press any key to exit...
pause > nul

REM Deactivate virtual environment
call venv\Scripts\deactivate.bat

exit /b 0
