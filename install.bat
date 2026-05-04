@echo off
echo ============================================
echo   New Viransh Wine Shop - Bill Calculator
echo   Installation Script
echo ============================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    py --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: Python is not installed or not in PATH.
        echo Please install Python 3.8+ from https://python.org
        echo Make sure to check "Add Python to PATH" during installation.
        pause
        exit /b 1
    )
)

echo [1/3] Python found.
echo.

:: Install dependencies
echo [2/3] Installing required packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    py -m pip install -r requirements.txt
)
echo.

:: Create desktop shortcut
echo [3/3] Creating desktop shortcut...
set SCRIPT_DIR=%~dp0
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Bill Calculator.lnk')); $s.TargetPath = 'pythonw'; $s.Arguments = '\"%SCRIPT_DIR%bill_calculator.py\"'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.Description = 'New Viransh Wine Shop - Bill Calculator'; $s.Save()"
echo.

echo ============================================
echo   Installation Complete!
echo ============================================
echo.
echo You can now:
echo   1. Double-click "Bill Calculator" shortcut on Desktop
echo   2. Or run: py bill_calculator.py
echo.
pause
