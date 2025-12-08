#!/bin/bash
# build_app.sh - macOS shell script to create .app bundle using PyInstaller

set -e  # Exit on any error

echo "🚀 Building Bank Schedule Sanitizer .app bundle for macOS..."
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

# Create the .app bundle
echo "📦 Creating macOS .app bundle with PyInstaller..."
pyinstaller --windowed \
    --name "Bank Schedule Sanitizer" \
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
    echo "📁 App bundle location: dist/Bank Schedule Sanitizer.app"
    echo
    echo "🎯 Next steps:"
    echo "   1. Test the app: open 'dist/Bank Schedule Sanitizer.app'"
    echo "   2. Copy the .app to Applications folder if desired"
    echo "   3. Distribute the .app bundle to end users"
    echo
    echo "📝 Note: This creates a proper macOS .app bundle that users can"
    echo "         double-click to launch from Finder."
    echo
    echo "💡 Tip: You can also drag the .app to the Dock for easy access"
    echo
else
    echo
    echo "❌ Build failed. Check the error messages above."
    echo
    exit 1
fi

echo "🔧 Build script completed."
