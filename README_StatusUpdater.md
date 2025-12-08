# Status Column Updater

A standalone Python script to update Status columns in Excel files with "LET" values for qualifying units.

## Features

- 🔍 **Smart Column Detection**: Automatically finds Status, Tenant Name, Start Date, and Expiry Date columns
- 🏠 **Unit Qualification**: Sets Status to "LET" for units that meet all criteria:
  - Has unit data (non-empty row)
  - Tenant Name is not vacant/null
  - Start Date is present
  - Expiry Date is present
- 💾 **Automatic Backup**: Creates backup files before making changes
- 🔒 **File Lock Detection**: Checks if files are open in other applications
- 📊 **Progress Reporting**: Shows detailed progress and statistics

## Usage

### Basic Usage (Default File)
```bash
python3 status_column_updater.py
```

### Custom File Path
```bash
python3 status_column_updater.py "/path/to/your/excel/file.xlsx"
```

### Make Executable (Optional)
```bash
chmod +x status_column_updater.py
./status_column_updater.py
```

## Default Target

The script is pre-configured to process:
```
data/30 November 2025 Bank Schedule [updated].xlsx
```

## Column Detection

The script automatically searches for columns with these names (case-insensitive):

- **Status**: "Status"
- **Tenant Name**: "Tenant Name"  
- **Start Date**: "Start Date", "Lease Start", "Start", "Commencement"
- **Expiry Date**: "Expiry Date", "End Date", "Lease End", "Expiry", "Termination"

## Status Criteria

A unit gets "LET" status if:
1. ✅ Row contains unit data
2. ✅ Tenant Name is not vacant (not empty, "VACANT", "TO LET", etc.)
3. ✅ Start Date has a valid date value
4. ✅ Expiry Date has a valid date value

## Output

The script provides:
- 📋 Backup file creation confirmation
- 📍 Column mapping details  
- 🔄 Row-by-row processing updates
- 📊 Final statistics (rows processed, updates made)
- ✅ Success confirmation with file save

## Example Output

```
=== Status Column Updater ===
Processing file: 30 November 2025 Bank Schedule [updated].xlsx
📋 Backup created: 30 November 2025 Bank Schedule [updated]_BACKUP_StatusUpdate.xlsx
📂 Opening workbook...
📋 Working with sheet: 'Bank Schedule'
🔍 Locating columns...
✅ Status column found: AE
📍 Column mapping:
   Status: AE
   Tenant Name: K
   Start Date: L
   Expiry Date: M

🔄 Processing rows...
   ✅ Row 5: Set status to 'LET'
   ✅ Row 7: Set status to 'LET'
   ...

📊 Processing complete:
   Rows processed: 965
   Updates made: 153
💾 Saving changes...
✅ File saved successfully!

🎉 Status column update completed successfully!
```

## Requirements

- Python 3.6+
- openpyxl
- pandas

## Notes

- The script starts processing from row 3 to skip headers
- Backup files are created with `_BACKUP_StatusUpdate.xlsx` suffix
- The script will not run if the target file is open in Excel
- Column AE is used as fallback if "Status" column is not found by name
