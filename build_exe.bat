@echo off
echo ============================================
echo   Building Standalone Executable (.exe)
echo ============================================
echo.
echo This will create a single .exe file that can
echo run without Python installed on the target PC.
echo.

:: Install PyInstaller if not present
echo [1/3] Installing PyInstaller...
pip install pyinstaller
echo.

:: Build the executable
echo [2/3] Building executable...
pyinstaller --onefile --windowed --name "BillCalculator" --add-data "items_data.xlsx;." bill_calculator.py
echo.

:: Copy data file to dist folder
echo [3/3] Copying data files...
if exist "dist\BillCalculator.exe" (
    copy "items_data.xlsx" "dist\" >nul 2>&1
    echo.
    echo ============================================
    echo   Build Complete!
    echo ============================================
    echo.
    echo Your standalone installer is ready at:
    echo   dist\BillCalculator.exe
    echo.
    echo To install on a new PC:
    echo   1. Copy the "dist" folder to the target PC
    echo   2. Double-click BillCalculator.exe
    echo   3. Make sure items_data.xlsx is in the same folder
    echo.
) else (
    echo ERROR: Build failed. Check the output above.
)
pause
