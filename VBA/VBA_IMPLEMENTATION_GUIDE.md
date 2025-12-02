# VBA Implementation Guide

## Overview
The VBA version provides the core functionality of the Python script directly in Excel:
- ✅ Creates Units sheet (filtered unit data)
- ✅ Creates Buildings sheet with SUMIF formulas
- ✅ Auto-runs on workbook open (optional)
- ⚠️ Cap Valn mapping needs manual setup (see below)

## Implementation Steps

### 1. Enable VBA in Excel
1. File → Options → Trust Center → Trust Center Settings
2. Enable "Trust access to the VBA project object model"
3. Enable macros for this workbook

### 2. Import the VBA Code
1. Press `Alt + F11` to open VBA Editor
2. Right-click on your workbook in Project Explorer
3. Insert → Module
4. Copy the code from `BankScheduleSanitizerVBA.bas`
5. Save the workbook as `.xlsm` (macro-enabled)

### 3. Running the Macro

#### Option A: Manual Run
- Press `Alt + F8`
- Select "ProcessBankSchedule"
- Click "Run"

#### Option B: Auto-run on Open
- The code includes `Workbook_Open()` event
- Will prompt user when file opens
- Can be customized or disabled

### 4. Setting Up Auto-Run
To make it run automatically on open:
1. In VBA Editor, double-click "ThisWorkbook"
2. Add this code:
```vba
Private Sub Workbook_Open()
    Call ProcessBankSchedule
End Sub
```

## Key Differences from Python Version

### ✅ **WHAT WORKS THE SAME**
- Units sheet creation with filtered data
- Buildings sheet with unique building list
- SUMIF formulas for aggregation
- Column detection and mapping
- Error handling and validation

### ⚠️ **MANUAL SETUP REQUIRED**

#### Cap Valn Mapping
The complex Cap Valn mapping from Python needs manual setup:

1. **After running the macro**, you'll see "Cap Valn lookup needed" in column F
2. **Replace with direct references** like:
   ```
   ='Bank Schedule'!AB7    (for first building)
   ='Bank Schedule'!AB8    (for second building)
   etc.
   ```
3. **Or create a lookup table** in a separate area with building names and Cap Valn references

## Advantages of VBA Version

1. **No External Dependencies** - Runs entirely within Excel
2. **Instant Execution** - No file saving/loading
3. **Live Updates** - Formulas update automatically
4. **User Friendly** - Familiar Excel environment
5. **Embedded Solution** - Travels with the workbook

## Limitations

1. **Cap Valn Complexity** - The sophisticated building name matching needs manual setup
2. **Performance** - May be slower with very large datasets
3. **Error Handling** - Less robust than Python version
4. **Debugging** - VBA debugging is more limited

## Recommended Workflow

### For Regular Use:
1. Use VBA version for day-to-day processing
2. Manual Cap Valn setup once per template
3. Auto-run on workbook open

### For Complex Cases:
1. Use Python version for initial setup
2. Copy Cap Valn formulas to VBA version
3. Switch to VBA for ongoing updates

## Performance Notes

- **Fast**: Units/Buildings sheet creation
- **Medium**: Formula calculations with large datasets  
- **Manual**: Cap Valn mapping (one-time setup)

The VBA version gives you 90% of the functionality with 100% Excel integration!
