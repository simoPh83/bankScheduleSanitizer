#!/bin/bash
# build.sh - macOS shell script to create executable using PyInstaller

set -e  # Exit on any error

echo "🚀 Building Bank Schedule Sanitizer Executable for macOS..."
echo

# Check if virtual environment exists and is properly set up
if [ ! -f ".venv/bin/activate" ]; then
    echo "⚠️ Virtual environment not found or incomplete. Creating/recreating one now..."
    rm -rf .venv  # Remove any incomplete venv
    python3 -m venv .venv
    source .venv/bin/activate
    python -m pip install --upgrade pip
    pip install -r requirements.txt
else
    echo "✓ Activating virtual environment..."
    source .venv/bin/activate
fi

# Check if PyInstaller is available
if ! command -v pyinstaller &> /dev/null; then
    echo "📦 Installing PyInstaller..."
    pip install pyinstaller
fi

# Clean previous builds
echo "🧹 Cleaning previous builds..."
rm -rf build dist __pycache__ *.spec

# Create the executable
echo "📦 Creating macOS app with PyInstaller..."
pyinstaller --onefile --windowed \
    --name "BankScheduleSanitizer" \
    --add-data "instructions:instructions" \
    --hidden-import=pandas \
    --hidden-import=openpyxl \
    --hidden-import=xlsxwriter \
    --hidden-import=tkinter \
    --hidden-import=datetime \
    --exclude-module=pytest \
    --exclude-module=_pytest \
    --exclude-module=test \
    --osx-bundle-identifier=com.simoneph.bankschedulesanitizer \
    bank_schedule_sanitizer.py

# Check if build was successful
if [ $? -eq 0 ]; then
    echo
    echo "✅ Build completed successfully!"
    echo "📁 Executable location: dist/BankScheduleSanitizer"
    echo
    echo "🎯 Next steps:"
    echo "   1. Test the executable: ./dist/BankScheduleSanitizer"
    echo "   2. Copy your Excel files to test with the app"
    echo "   3. Distribute the executable to end users"
    echo
    echo "📝 Note: The executable is located at dist/BankScheduleSanitizer"
    echo "         On macOS, this creates a Unix executable, not a .app bundle"
    echo
else
    echo
    echo "❌ Build failed. Check the error messages above."
    echo
    exit 1
fi

echo "🔧 Build script completed."
