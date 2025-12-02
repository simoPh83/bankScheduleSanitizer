# VBA Troubleshooting Guide

## 🚨 **IF EXCEL FREEZES AGAIN**

### IMMEDIATE STEPS:
1. **Don't panic!** Press `Ctrl + Break` to stop the macro
2. If that doesn't work, use Task Manager to force-close Excel
3. Reopen Excel and **disable macros** when prompted
4. Open the Immediate Window (`Ctrl + G`) to see debug messages

---

## 🔧 **UPDATED SAFE VERSION FEATURES**

The new VBA code includes:

### ✅ **SAFETY MECHANISMS**
- **Row/Column limits**: Max 50,000 rows, 100 columns
- **Progress updates**: Status bar shows current operation
- **DoEvents**: Allows Excel to respond during processing
- **Error recovery**: Each section has individual error handling
- **Memory management**: Careful copy/paste operations

### 🐛 **DEBUG FEATURES**
- **Debug messages**: Check Immediate Window (`Ctrl + G`)
- **Progress tracking**: See exactly where processing is
- **Error details**: Specific error messages with context
- **Validation**: Checks data size and sheet existence

### 📊 **PERFORMANCE OPTIMIZATIONS**
- **Bulk operations**: Copies ranges instead of individual cells
- **Calculation control**: Turns off auto-calculation during processing
- **Screen updates**: Disabled during processing for speed
- **Event handling**: Temporarily disabled to prevent conflicts

---

## 🎯 **HOW TO TEST SAFELY**

### Step 1: Enable Debug Mode
```vba
Const DEBUG_MODE = True ' Already set to True in the code
```

### Step 2: Test with Small Data First
1. Try with a small section of your Bank Schedule (50-100 rows)
2. Watch the Immediate Window for progress messages
3. If successful, try with full data

### Step 3: Monitor Progress
- Status bar shows current operation
- Immediate Window shows detailed progress
- Can press `Ctrl + Break` to stop if needed

---

## 🔍 **COMMON ISSUES & SOLUTIONS**

### **Issue: "Bank Schedule sheet not found"**
**Solution**: Ensure your source sheet is named exactly "Bank Schedule" (case-sensitive)

### **Issue: "Could not find required columns"**
**Solution**: 
- Check that row 3 contains headers
- Look for columns containing "Unit Demise" and "Property"
- Check Immediate Window for found column numbers

### **Issue: "Data too large"**
**Solution**: The code limits to 50,000 rows for safety. If you need more:
```vba
Const MAX_ROWS = 100000 ' Increase this number
```

### **Issue: Formula errors in Buildings sheet**
**Solution**: 
- Check if column names match exactly
- Look for special characters in building names
- Yellow cells show Cap Valn needs manual setup

### **Issue: Performance still slow**
**Solution**: 
- Temporarily set `DEBUG_MODE = False` for faster processing
- Close other Excel workbooks
- Ensure sufficient RAM available

---

## 📋 **TESTING CHECKLIST**

Before running on production data:

1. **✅ Backup your file**
2. **✅ Enable debug mode**
3. **✅ Test with small dataset first**
4. **✅ Watch Immediate Window (`Ctrl + G`)**
5. **✅ Verify column names in row 3**
6. **✅ Check data doesn't exceed 50,000 rows**

---

## 🚀 **IF EVERYTHING WORKS**

You'll get:
- **Units sheet**: Clean unit data with formulas
- **Buildings sheet**: Summary with SUMIF formulas
- **Progress messages**: Clear feedback on what's happening
- **Cap Valn placeholders**: Ready for manual setup

The yellow cells in Buildings sheet show where to add Cap Valn references like:
```
='Bank Schedule'!AB7
='Bank Schedule'!AB8
```

---

## 🆘 **EMERGENCY FALLBACK**

If the full version still has issues, use the **Simple version** instead:
- File: `BankScheduleSanitizer_Simple.bas`
- More basic but very stable
- Same core functionality, less bells and whistles

---

## 📞 **DEBUG OUTPUT EXAMPLE**

In the Immediate Window, you should see:
```
Starting ProcessBankSchedule at 14:30:15
Validating environment...
Environment validated. Rows: 1250, Cols: 35
Starting CreateUnitsSheet...
Data bounds - LastRow: 1250, LastCol: 35
Column positions - Building: 5, UnitDemise: 12, Property: 15
Copying headers...
Copied 847 unit rows
Units sheet created successfully
```

This shows everything is working correctly!
