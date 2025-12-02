@echo off
REM build.bat - Windows batch script to create executable using PyInstaller

echo 🚀 Building Bank Schedule Sanitizer Executable...
echo.

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo ⚠️ Virtual environment not found. Creating one now...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    python -m pip install --upgrade pip
    pip install -r requirements.txt
) else (
    echo ✓ Activating virtual environment...
    call .venv\Scripts\activate.bat
)

REM Clean previous builds
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM Create the executable
echo 📦 Creating executable with PyInstaller...
.venv\Scripts\pyinstaller.exe --onefile --windowed ^
    --name "BankScheduleSanitizer" ^
    --add-data "instructions;instructions" ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=xlsxwriter ^
    --exclude-module=pytest ^
    --exclude-module=_pytest ^
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
