@echo off
echo ===============================================================
echo               PDF Extractor - Uninstaller
echo ===============================================================
echo.

echo This will remove the application files and optionally uninstall packages.
set /p confirm=Are you sure you want to proceed? (Y/N): 
if /i "%confirm%" neq "Y" (
    echo Uninstallation cancelled.
    pause
    exit /b 0
)

echo.
echo Removing generated files...

if exist "tax_invoice_data.xlsx" (
    del /q "tax_invoice_data.xlsx"
    echo - Removed output Excel file
)

if exist "pdf_extraction.log" (
    del /q "pdf_extraction.log"
    echo - Removed log file
)

:: Remove any text debug files
del /q "*_text.txt" 2>nul
echo - Removed debug text files (if any)

:: Ask about removing virtual environment
if exist "venv" (
    set /p remove_env=Do you want to remove the virtual environment? (Y/N): 
    if /i "%remove_env%"=="Y" (
        echo Removing virtual environment...
        rmdir /s /q venv
        echo - Virtual environment removed
    )
)

:: Ask about uninstalling packages
set /p uninstall_packages=Do you want to uninstall the Python packages (pdfplumber, pandas, openpyxl)? (Y/N): 
if /i "%uninstall_packages%"=="Y" (
    echo Uninstalling packages...
    pip uninstall -y pdfplumber pandas openpyxl
    echo - Packages uninstalled
)

echo.
echo Uninstallation completed successfully!
echo.
echo Press any key to exit...
pause > nul
exit /b 0
