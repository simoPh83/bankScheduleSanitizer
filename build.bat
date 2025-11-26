@echo off
REM build.bat - Windows batch script to create executable using PyInstaller

echo 🚀 Building Bank Schedule Sanitizer Executable...
echo.

REM Check if virtual environment is activated
if "%VIRTUAL_ENV%"=="" (
    echo ⚠️ Warning: Virtual environment not detected. Please activate your venv first:
    echo    python -m venv .venv
    echo    .venv\Scripts\activate
    echo    pip install -r requirements.txt
    echo.
)

REM Create the executable
echo 📦 Creating executable with PyInstaller...
pyinstaller --onefile --windowed ^
    --name "BankScheduleSanitizer" ^
    --add-data "instructions;instructions" ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=xlsxwriter ^
    bank_schedule_sanitizer.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo ✅ Build completed successfully!
    echo 📁 Executable location: dist\BankScheduleSanitizer.exe
    echo.
    echo 🎯 Next steps:
    echo    1. Test the executable: dist\BankScheduleSanitizer.exe
    echo    2. Copy your Excel files to the same folder as the executable
    echo    3. Distribute the executable to end users
    echo.
) else (
    echo.
    echo ❌ Build failed. Check the error messages above.
    echo.
)

pause
