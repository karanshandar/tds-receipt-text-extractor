@echo off
echo ===============================================================
echo                PDF Data Extractor
echo ===============================================================
echo.

REM Check if virtual environment exists
if not exist "venv" (
    echo Virtual environment not found.
    echo Please run setup.bat first to set up the environment.
    pause
    exit /b 1
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Failed to activate virtual environment.
    pause
    exit /b 1
)

REM Run the Python script
echo.
echo Running PDF extraction...
echo.

REM Check for debug flag
set debug_flag=
if "%~1"=="--debug" set debug_flag=--debug
if "%~1"=="-d" set debug_flag=--debug

REM Run the script with optional debug flag
python pdf_extractor.py %debug_flag%

REM Check if extraction was successful and open the Excel file
if exist tax_invoice_data.xlsx (
    echo.
    echo Extraction completed successfully!
    echo Results saved to tax_invoice_data.xlsx
    
    REM Ask if user wants to open the Excel file
    set /p open_excel=Do you want to open the Excel file now? (Y/N): 
    if /i "%open_excel%"=="Y" (
        start tax_invoice_data.xlsx
    )
) else (
    echo.
    echo Extraction completed, but no output file was created.
    echo Check the log file for details.
)

echo.
echo Press any key to exit...
pause > nul

REM Deactivate virtual environment
call venv\Scripts\deactivate.bat

exit /b 0
