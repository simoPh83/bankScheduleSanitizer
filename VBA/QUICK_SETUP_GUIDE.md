# Quick VBA Setup Guide

## 🚀 **5-MINUTE SETUP**

### Step 1: Enable Macros
1. Open your Excel file with Bank Schedule data
2. File → Options → Trust Center → Trust Center Settings
3. Check "Enable all macros" (for testing)
4. Check "Trust access to the VBA project object model"

### Step 2: Import VBA Code
1. Press `Alt + F11` (opens VBA Editor)
2. Right-click your workbook name → Insert → Module
3. Copy code from `BankScheduleSanitizer_Simple.bas`
4. Paste into the module window
5. Save workbook as `.xlsm` format

### Step 3: Run the Macro
1. Press `Alt + F8`
2. Select "ProcessBankScheduleSimple" 
3. Click Run

**That's it!** You'll get Units and Buildings sheets with formulas.

---

## 🔧 **AUTO-RUN ON OPEN** (Optional)

To run automatically when opening the file:

1. In VBA Editor, double-click "ThisWorkbook"
2. Add this code:
```vba
Private Sub Workbook_Open()
    ProcessBankScheduleSimple
End Sub
```

---

## ⚙️ **CAP VALN SETUP** (One-time)

The macro creates placeholder "Manual setup needed" for Cap Valn values. Replace with:

### Option A: Direct References
```
='Bank Schedule'!AB7
='Bank Schedule'!AB8
='Bank Schedule'!AB9
```

### Option B: Lookup Formula
```
=INDEX('Bank Schedule'!AB:AB,MATCH(A2,'Bank Schedule'!A:A,0))
```

---

## ✅ **WHAT YOU GET**

- **Units Sheet**: All unit data (no building summary rows)
- **Buildings Sheet**: Unique buildings with SUMIF formulas for:
  - Net Area
  - Rent PA
  - 2023 ERV  
  - 2024 ERV
  - Cap Valn (needs setup)

---

## 🆚 **VBA vs Python Comparison**

| Feature | VBA Version | Python Version |
|---------|-------------|----------------|
| **Setup** | 5 minutes | Install Python + dependencies |
| **Speed** | Fast | Very fast |
| **File handling** | Modifies existing | Creates new files |
| **Cap Valn** | Manual setup | Automatic matching |
| **Updates** | Live formulas | Run script again |
| **Portability** | Embedded in Excel | Separate executable |
| **User friendliness** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ |

---

## 🎯 **RECOMMENDATION**

**Use VBA version if:**
- You want everything in Excel
- You're comfortable with one-time Cap Valn setup
- You want auto-run on file open
- You prefer live updating formulas

**Use Python version if:**
- You have complex building name variations
- You want fully automated Cap Valn mapping  
- You process many different files
- You want robust error handling

Both versions give you the same core functionality!
