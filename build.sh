#!/bin/bash
# build.sh - Script to create Windows executable using PyInstaller
# Run this on Windows after setting up the environment

echo "🚀 Building Bank Schedule Sanitizer Executable..."
echo ""

# Check if virtual environment is activated
if [[ "$VIRTUAL_ENV" == "" ]]; then
    echo "⚠️  Warning: Virtual environment not detected. Please activate your venv first:"
    echo "   python -m venv .venv"
    echo "   .venv\\Scripts\\activate  (on Windows)"
    echo "   pip install -r requirements.txt"
    echo ""
fi

# Create the executable
echo "📦 Creating executable with PyInstaller..."
pyinstaller --onefile --windowed \
    --name "BankScheduleSanitizer" \
    --icon="data/app_icon.ico" \
    --add-data "instructions;instructions" \
    --hidden-import=pandas \
    --hidden-import=openpyxl \
    --hidden-import=xlsxwriter \
    bank_schedule_sanitizer.py

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Build completed successfully!"
    echo "📁 Executable location: dist/BankScheduleSanitizer.exe"
    echo ""
    echo "🎯 Next steps:"
    echo "   1. Test the executable: dist/BankScheduleSanitizer.exe"
    echo "   2. Copy your Excel files to the same folder as the executable"
    echo "   3. Distribute the executable to end users"
    echo ""
else
    echo ""
    echo "❌ Build failed. Check the error messages above."
    echo ""
fi
