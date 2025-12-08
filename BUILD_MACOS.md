# Build Instructions for macOS

This directory contains build scripts for creating distributable executables of the Bank Schedule Sanitizer application on macOS.

## Available Build Scripts

### 1. `build.sh` - Single Executable ⚠️ **Deprecated on macOS**
Creates a single executable file that can be run from the terminal.

```bash
./build.sh
```

**Output**: `dist/BankScheduleSanitizer` (Unix executable)
**Usage**: Run via terminal: `./dist/BankScheduleSanitizer`
**Note**: ⚠️ PyInstaller 7.0+ will no longer support `--onefile` + `--windowed` on macOS

### 2. `build_app.sh` - macOS App Bundle ✅ **Recommended**
Creates a proper macOS .app bundle that users can double-click.

```bash
./build_app.sh
```

**Output**: `dist/Bank Schedule Sanitizer.app` (macOS app bundle)
**Usage**: Double-click the .app file or run `open "dist/Bank Schedule Sanitizer.app"`

## Prerequisites

- Python 3.7 or later
- macOS 10.12 or later
- All dependencies listed in `requirements.txt`

## Build Process

Both scripts will automatically:
1. Create a virtual environment if one doesn't exist
2. Install required dependencies including PyInstaller
3. Clean previous builds
4. Create the executable/app bundle
5. Provide next steps for testing and distribution

## Distribution

### Single Executable (`build.sh`)
- Distribute the `BankScheduleSanitizer` file
- Users run it from terminal: `./BankScheduleSanitizer`
- Smaller file size, but requires terminal to launch

### App Bundle (`build_app.sh`) - **Recommended**
- Distribute the entire `Bank Schedule Sanitizer.app` bundle
- Users can double-click to launch like any Mac app
- Can be moved to Applications folder
- More user-friendly for non-technical users

## Troubleshooting

If you get permission errors, make sure the scripts are executable:
```bash
chmod +x build.sh
chmod +x build_app.sh
```

For more detailed build options or issues, check the PyInstaller documentation.
