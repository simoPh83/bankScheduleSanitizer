# Windows Deployment Guide

## 📦 Creating Windows Executable

### Prerequisites
1. **Python 3.9+** installed on Windows PC
2. **Git** (to clone/transfer the project)

### Setup Steps

#### 1. Transfer Project Files
```bash
# Option A: Git clone (if using Git)
git clone <repository-url>
cd bankScheduleSanitizer

# Option B: Copy folder directly to Windows PC
# Just copy the entire bankScheduleSanitizer folder
```

#### 2. Create Virtual Environment
```cmd
python -m venv .venv
.venv\Scripts\activate
```

#### 3. Install Dependencies
```cmd
pip install -r requirements.txt
```

#### 4. Build Executable
```cmd
# Run the build script
build.bat

# Or manually run PyInstaller:
pyinstaller --onefile --windowed --name "BankScheduleSanitizer" bank_schedule_sanitizer.py
```

### 🎯 Output
- **Executable**: `dist/BankScheduleSanitizer.exe`
- **Size**: ~50-100MB (includes Python runtime and all dependencies)
- **Standalone**: No Python installation required on target machines

### 📋 Distribution
1. Copy `BankScheduleSanitizer.exe` to target machines
2. Include sample Excel files if needed
3. No additional installation required

### 🔧 Troubleshooting

#### Build Issues
- **Missing modules**: Add `--hidden-import=module_name` to PyInstaller command
- **File not found**: Ensure all files are in the project directory
- **Permission errors**: Run Command Prompt as Administrator

#### Runtime Issues
- **Antivirus blocking**: Add executable to antivirus whitelist
- **DLL errors**: Ensure target Windows has Visual C++ Redistributable

### 📁 Project Structure After Build
```
bankScheduleSanitizer/
├── dist/
│   └── BankScheduleSanitizer.exe    # ← Distribute this file
├── build/                           # ← Temporary build files
├── BankScheduleSanitizer.spec       # ← PyInstaller spec file
├── data/                            # ← Excel files
├── instructions/                    # ← Documentation
└── [source files...]               # ← Development files
```

### 🚀 Advanced Options

#### Custom Icon
```cmd
pyinstaller --onefile --windowed --icon="icon.ico" bank_schedule_sanitizer.py
```

#### Include Additional Files
```cmd
pyinstaller --onefile --add-data "templates;templates" bank_schedule_sanitizer.py
```

#### Debug Version (shows console)
```cmd
pyinstaller --onefile --console bank_schedule_sanitizer.py
```
