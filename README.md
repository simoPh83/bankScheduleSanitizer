# Bank Schedule Sanitizer

A GUI application to process and sanitize Excel bank schedule files.

## Features

- **File Selection**: Easy-to-use interface for selecting input Excel files
- **Output Location**: Choose where to save the processed file
- **Error Logging**: Built-in text box for status messages and error reporting
- **Sheet Validation**: Checks for the presence of "Bank Schedule" sheet
- **User-Friendly**: Simple, intuitive interface with clear status messages

## Setup Instructions

### 1. Create Virtual Environment
```bash
python -m venv .venv
```

### 2. Activate Virtual Environment

**macOS/Linux:**
```bash
source .venv/bin/activate
```

**Windows:**
```cmd
.venv\Scripts\activate
```

### 3. Install Required Packages
```bash
pip install -r requirements.txt
```

### 4. Run the Application
```bash
python bank_schedule_sanitizer.py
```

## Usage

1. **Launch the application** by running the Python script
2. **Select an input file** by clicking the "Browse..." button next to "Select Excel file to process"
3. **Choose output location** by clicking the "Browse..." button next to "Save sanitized file as"
4. **Click "Sanitize"** to process the file
5. **Monitor the status** in the text box at the bottom for any messages or errors

## Current Functionality

Currently, the application:
- Validates that the input file is a readable Excel file
- Checks for the presence of the "Bank Schedule" sheet
- Duplicates the file to the specified output location
- Provides detailed logging of the process

## Future Enhancements

The sanitization logic will be implemented based on specific requirements for processing the "Bank Schedule" sheet data.

## Requirements

- Python 3.7+
- pandas
- openpyxl
- xlsxwriter
- tkinter (usually included with Python)

## File Structure

```
bankScheduleSanitizer/
├── bank_schedule_sanitizer.py  # Main application
├── requirements.txt            # Python dependencies
├── README.md                  # This file
├── data/                      # Sample data files
└── instructions/              # Project instructions
```
