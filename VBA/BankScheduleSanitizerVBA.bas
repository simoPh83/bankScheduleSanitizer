' BankScheduleSanitizerVBA.bas - SAFE VERSION
' VBA version of the Bank Schedule Sanitizer with robust error handling
' Converts Bank Schedule data into Units and Buildings sheets with formulas

Option Explicit

' Global constants for safety
Const MAX_ROWS = 50000 ' Safety limit for processing
Const MAX_COLS = 100   ' Safety limit for columns
Const DEBUG_MODE = True ' Set to False to disable debug messages

' Main entry point - can be triggered on workbook open or manually
Sub ProcessBankSchedule()
    On Error GoTo ErrorHandler
    
    ' Safety: Turn off events and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim startTime As Double
    startTime = Timer
    
    ' Debug message
    If DEBUG_MODE Then Debug.Print "Starting ProcessBankSchedule at " & Format(Now, "hh:mm:ss")
    
    ' Validate environment
    If Not ValidateEnvironment() Then Exit Sub
    
    ' Show progress
    StatusBarMessage "Processing Bank Schedule - Creating Units sheet..."
    
    ' Run the processing with individual error handling
    If CreateUnitsSheetSafe() Then
        StatusBarMessage "Processing Bank Schedule - Creating Buildings sheet..."
        Call CreateBuildingsSheetSafe
    End If
    
    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Dim processingTime As Double
    processingTime = Timer - startTime
    
    If DEBUG_MODE Then Debug.Print "Processing completed in " & Format(processingTime, "0.0") & " seconds"
    
    MsgBox "File processed successfully!" & vbCrLf & _
           "Processing time: " & Format(processingTime, "0.0") & " seconds" & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for debug details", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' Restore settings on error
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Dim errorMsg As String
    errorMsg = "Error in ProcessBankSchedule: " & Err.Description & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "At: " & Format(Now, "hh:mm:ss")
               
    If DEBUG_MODE Then Debug.Print errorMsg
    MsgBox errorMsg, vbCritical
End Sub

' Validate environment before processing
Function ValidateEnvironment() As Boolean
    On Error GoTo ValidationError
    
    If DEBUG_MODE Then Debug.Print "Validating environment..."
    
    ' Check if Bank Schedule sheet exists
    If Not SheetExists("Bank Schedule") Then
        MsgBox "Bank Schedule sheet not found!" & vbCrLf & _
               "Please ensure your source sheet is named exactly 'Bank Schedule'", vbCritical
        ValidateEnvironment = False
        Exit Function
    End If
    
    ' Check if Bank Schedule has data
    Dim bankWS As Worksheet
    Set bankWS = Worksheets("Bank Schedule")
    
    ' Check multiple locations for data instead of just (1,1)
    Dim hasData As Boolean
    hasData = False
    
    ' Check first few rows and columns for any data
    Dim checkRow As Long, checkCol As Long
    For checkRow = 1 To 10
        For checkCol = 1 To 10
            If Len(Trim(bankWS.Cells(checkRow, checkCol).Value)) > 0 Then
                hasData = True
                Exit For
            End If
        Next checkCol
        If hasData Then Exit For
    Next checkRow
    
    If Not hasData Then
        MsgBox "Bank Schedule sheet appears to be empty!" & vbCrLf & _
               "Checked first 10 rows and columns for data.", vbCritical
        ValidateEnvironment = False
        Exit Function
    End If
    
    If DEBUG_MODE Then Debug.Print "Data found in Bank Schedule sheet"
    
    ' Check for reasonable data size with better detection
    Dim lastRow As Long, lastCol As Long
    
    ' Find the actual last row with data (check multiple columns)
    lastRow = 0
    For checkCol = 1 To 20 ' Check first 20 columns
        Dim colLastRow As Long
        colLastRow = bankWS.Cells(bankWS.Rows.Count, checkCol).End(xlUp).Row
        If colLastRow > lastRow Then lastRow = colLastRow
    Next checkCol
    
    ' Find the actual last column with data (check row 3 for headers)
    lastCol = bankWS.Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' If row 3 has no data, check other rows
    If lastCol <= 1 Then
        For checkRow = 1 To 10
            Dim rowLastCol As Long
            rowLastCol = bankWS.Cells(checkRow, Columns.Count).End(xlToLeft).Column
            If rowLastCol > lastCol Then lastCol = rowLastCol
        Next checkRow
    End If
    
    If lastRow > MAX_ROWS Then
        MsgBox "Data too large! Found " & lastRow & " rows. Maximum supported: " & MAX_ROWS, vbCritical
        ValidateEnvironment = False
        Exit Function
    End If
    
    If lastCol > MAX_COLS Then
        MsgBox "Too many columns! Found " & lastCol & " columns. Maximum supported: " & MAX_COLS, vbCritical
        ValidateEnvironment = False
        Exit Function
    End If
    
    If DEBUG_MODE Then Debug.Print "Environment validated. Rows: " & lastRow & ", Cols: " & lastCol
    
    ' Additional debug info about the sheet structure
    If DEBUG_MODE Then
        Debug.Print "Sample data check:"
        Debug.Print "  Cell A1: '" & bankWS.Cells(1, 1).Value & "'"
        Debug.Print "  Cell A3: '" & bankWS.Cells(3, 1).Value & "'"
        Debug.Print "  Cell B3: '" & bankWS.Cells(3, 2).Value & "'"
        Debug.Print "  Cell C3: '" & bankWS.Cells(3, 3).Value & "'"
    End If
    
    ValidateEnvironment = True
    Exit Function
    
ValidationError:
    MsgBox "Error during environment validation: " & Err.Description, vbCritical
    ValidateEnvironment = False
End Function

' Create Units sheet from Bank Schedule data - SAFE VERSION
Function CreateUnitsSheetSafe() As Boolean
    On Error GoTo UnitsError
    
    If DEBUG_MODE Then Debug.Print "Starting CreateUnitsSheet..."
    
    Dim bankWS As Worksheet
    Dim unitsWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, rowsProcessed As Long
    Dim buildingCol As Long, unitDemiseCol As Long, propertyCol As Long
    Dim checkCol As Long, checkRow As Long ' For data boundary detection
    
    Set bankWS = Worksheets("Bank Schedule")
    
    ' Find data boundaries with safety checks - use same logic as validation
    lastRow = 0
    For checkCol = 1 To 20 ' Check first 20 columns for real last row
        Dim colLastRow As Long
        colLastRow = bankWS.Cells(bankWS.Rows.Count, checkCol).End(xlUp).Row
        If colLastRow > lastRow Then lastRow = colLastRow
    Next checkCol
    
    lastCol = bankWS.Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' If row 3 has no data, check other rows for last column
    If lastCol <= 1 Then
        For checkRow = 1 To 10
            Dim rowLastCol As Long
            rowLastCol = bankWS.Cells(checkRow, Columns.Count).End(xlToLeft).Column
            If rowLastCol > lastCol Then lastCol = rowLastCol
        Next checkRow
    End If
    
    If DEBUG_MODE Then Debug.Print "Data bounds - LastRow: " & lastRow & ", LastCol: " & lastCol
    
    ' Safety checks - be more flexible about minimum rows
    If lastRow < 3 Then
        MsgBox "Not enough data rows in Bank Schedule" & vbCrLf & _
               "Found only " & lastRow & " rows. Need at least 3 rows (for headers + data).", vbCritical
        CreateUnitsSheetSafe = False
        Exit Function
    End If
    
    ' Additional check - make sure we have reasonable column count
    If lastCol < 3 Then
        MsgBox "Not enough columns found in Bank Schedule" & vbCrLf & _
               "Found only " & lastCol & " columns. This seems too few for a proper schedule.", vbCritical
        CreateUnitsSheetSafe = False
        Exit Function
    End If
    
    ' Delete existing Units sheet if it exists
    If SheetExists("Units") Then
        If DEBUG_MODE Then Debug.Print "Deleting existing Units sheet..."
        Application.DisplayAlerts = False
        Worksheets("Units").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new Units sheet
    If DEBUG_MODE Then Debug.Print "Creating new Units sheet..."
    Set unitsWS = Worksheets.Add
    unitsWS.Name = "Units"
    
    ' Find column positions with better error handling
    If DEBUG_MODE Then Debug.Print "Finding column positions..."
    buildingCol = FindColumnByHeaderSafe(bankWS, "Unit Type", 3)
    unitDemiseCol = FindColumnByHeaderSafe(bankWS, "Unit Demise", 3)
    propertyCol = FindColumnByHeaderSafe(bankWS, "Property", 3)
    
    If DEBUG_MODE Then Debug.Print "Column positions - Building: " & buildingCol & ", UnitDemise: " & unitDemiseCol & ", Property: " & propertyCol
    
    If unitDemiseCol = 0 Or propertyCol = 0 Then
        MsgBox "Could not find required columns in Bank Schedule" & vbCrLf & _
               "Looking for: Unit Demise, Property" & vbCrLf & _
               "Found UnitDemise: " & unitDemiseCol & ", Property: " & propertyCol, vbCritical
        CreateUnitsSheetSafe = False
        Exit Function
    End If
    
    ' Copy headers with progress tracking
    If DEBUG_MODE Then Debug.Print "Copying headers..."
    StatusBarMessage "Copying headers to Units sheet..."
    
    Dim targetCol As Long
    targetCol = 1
    
    For i = 1 To lastCol
        If i <> buildingCol Then ' Skip Unit Type column to avoid duplication
            bankWS.Cells(3, i).Copy
            unitsWS.Cells(1, targetCol).PasteSpecial xlPasteValues
            unitsWS.Cells(1, targetCol).PasteSpecial xlPasteFormats
            targetCol = targetCol + 1
        End If
        
        ' Progress update every 10 columns
        If i Mod 10 = 0 Then
            StatusBarMessage "Copying headers... " & i & "/" & lastCol
            DoEvents ' Allow Excel to respond
        End If
    Next i
    
    Application.CutCopyMode = False
    
    ' Copy unit data with progress tracking
    If DEBUG_MODE Then Debug.Print "Starting to copy unit data from row 4 to " & lastRow
    Dim unitsRow As Long
    unitsRow = 2
    rowsProcessed = 0
    
    For i = 4 To lastRow ' Start from row 4 (after headers)
        ' Progress update every 100 rows
        If rowsProcessed Mod 100 = 0 Then
            StatusBarMessage "Processing row " & i & "/" & lastRow & " (Found " & (unitsRow - 2) & " units)"
            DoEvents ' Allow Excel to respond
        End If
        
        ' Check if this row has unit data with better validation
        Dim hasUnitData As Boolean
        hasUnitData = (Len(Trim(bankWS.Cells(i, unitDemiseCol).Value)) > 0 And _
                      Len(Trim(bankWS.Cells(i, propertyCol).Value)) > 0)
        
        If hasUnitData Then
            ' Copy this row to Units sheet (excluding Unit Type column)
            targetCol = 1
            
            For j = 1 To lastCol
                If j <> buildingCol Then
                    On Error Resume Next ' Handle individual cell copy errors
                    bankWS.Cells(i, j).Copy
                    unitsWS.Cells(unitsRow, targetCol).PasteSpecial xlPasteAll
                    On Error GoTo UnitsError
                    targetCol = targetCol + 1
                End If
            Next j
            
            unitsRow = unitsRow + 1
        End If
        
        rowsProcessed = rowsProcessed + 1
        
        ' Safety break if too many units
        If unitsRow > MAX_ROWS Then
            MsgBox "Too many units found! Stopping at " & MAX_ROWS & " units for safety.", vbExclamation
            Exit For
        End If
    Next i
    
    Application.CutCopyMode = False
    
    If DEBUG_MODE Then Debug.Print "Copied " & (unitsRow - 2) & " unit rows"
    
    ' Format the Units sheet
    If DEBUG_MODE Then Debug.Print "Formatting Units sheet..."
    StatusBarMessage "Formatting Units sheet..."
    
    With unitsWS.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        On Error Resume Next ' AutoFilter might fail
        .AutoFilter
        On Error GoTo UnitsError
    End With
    
    ' Auto-resize columns safely
    On Error Resume Next
    unitsWS.Columns.AutoFit
    On Error GoTo UnitsError
    
    If DEBUG_MODE Then Debug.Print "Units sheet created successfully"
    CreateUnitsSheetSafe = True
    Exit Function
    
UnitsError:
    If DEBUG_MODE Then Debug.Print "Error in CreateUnitsSheet: " & Err.Description
    MsgBox "Error creating Units sheet: " & Err.Description & vbCrLf & _
           "Processed " & rowsProcessed & " rows before error", vbCritical
    CreateUnitsSheetSafe = False
End Function

' Create Buildings summary sheet with formulas - SAFE VERSION
Sub CreateBuildingsSheetSafe()
    On Error GoTo BuildingsError
    
    If DEBUG_MODE Then Debug.Print "Starting CreateBuildingsSheet..."
    
    Dim unitsWS As Worksheet
    Dim buildingsWS As Worksheet
    Dim buildingNames As Collection
    Dim propertyCol As Long
    
    ' Check if Units sheet exists
    If Not SheetExists("Units") Then
        MsgBox "Units sheet must be created first!", vbCritical
        Exit Sub
    End If
    
    Set unitsWS = Worksheets("Units")
    
    ' Delete existing Buildings sheet if it exists
    If SheetExists("Buildings") Then
        If DEBUG_MODE Then Debug.Print "Deleting existing Buildings sheet..."
        Application.DisplayAlerts = False
        Worksheets("Buildings").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new Buildings sheet
    If DEBUG_MODE Then Debug.Print "Creating new Buildings sheet..."
    Set buildingsWS = Worksheets.Add
    buildingsWS.Name = "Buildings"
    
    ' Find Property column in Units sheet
    propertyCol = FindColumnByHeaderSafe(unitsWS, "Property", 1)
    If propertyCol = 0 Then
        MsgBox "Could not find Property column in Units sheet", vbCritical
        Exit Sub
    End If
    
    If DEBUG_MODE Then Debug.Print "Property column found at: " & propertyCol
    
    ' Extract unique building names with progress
    StatusBarMessage "Extracting unique building names..."
    Set buildingNames = GetUniqueBuildingNamesSafe(unitsWS, propertyCol)
    
    If buildingNames.Count = 0 Then
        MsgBox "No building names found in Units sheet!", vbCritical
        Exit Sub
    End If
    
    If DEBUG_MODE Then Debug.Print "Found " & buildingNames.Count & " unique buildings"
    
    ' Create headers for Buildings sheet
    StatusBarMessage "Creating Buildings sheet headers..."
    Dim headers As Variant
    headers = Array("Building", "Net Area", "Rent PA (£)", "2023 ERV (£)", "2024 ERV (£)", "2024 Cap Valn. (£)")
    
    Dim i As Long
    For i = 0 To UBound(headers)
        buildingsWS.Cells(1, i + 1).Value = headers(i)
        buildingsWS.Cells(1, i + 1).Font.Bold = True
        buildingsWS.Cells(1, i + 1).Interior.Color = RGB(220, 220, 220)
    Next i
    
    ' Create building summary rows with formulas
    StatusBarMessage "Creating formulas for " & buildingNames.Count & " buildings..."
    Call CreateBuildingFormulasSafe(buildingsWS, unitsWS, buildingNames, propertyCol)
    
    ' Format the Buildings sheet
    If DEBUG_MODE Then Debug.Print "Formatting Buildings sheet..."
    StatusBarMessage "Formatting Buildings sheet..."
    
    With buildingsWS.Rows(1)
        .Font.Bold = True
        On Error Resume Next
        .AutoFilter
        On Error GoTo BuildingsError
    End With
    
    ' Auto-resize columns safely
    On Error Resume Next
    buildingsWS.Columns.AutoFit
    On Error GoTo BuildingsError
    
    If DEBUG_MODE Then Debug.Print "Buildings sheet created successfully"
    Exit Sub
    
BuildingsError:
    If DEBUG_MODE Then Debug.Print "Error in CreateBuildingsSheet: " & Err.Description
    MsgBox "Error creating Buildings sheet: " & Err.Description, vbCritical
End Sub

' Helper function to find column by header name - SAFE VERSION
Function FindColumnByHeaderSafe(ws As Worksheet, headerName As String, headerRow As Long) As Long
    On Error GoTo FindError
    
    If DEBUG_MODE Then Debug.Print "        FindColumnByHeaderSafe: Looking for '" & headerName & "' in row " & headerRow
    
    Dim i As Long
    Dim lastCol As Long
    Dim cellValue As String
    
    lastCol = ws.Cells(headerRow, Columns.Count).End(xlToLeft).Column
    If DEBUG_MODE Then Debug.Print "        FindColumnByHeaderSafe: Last column = " & lastCol
    
    ' Safety check
    If lastCol > MAX_COLS Then lastCol = MAX_COLS
    
    For i = 1 To lastCol
        cellValue = Trim(ws.Cells(headerRow, i).Value)
        If InStr(1, cellValue, headerName, vbTextCompare) > 0 Then
            If DEBUG_MODE Then Debug.Print "        FindColumnByHeaderSafe: Found '" & headerName & "' at column " & i
            FindColumnByHeaderSafe = i
            Exit Function
        End If
    Next i
    
    If DEBUG_MODE Then Debug.Print "        FindColumnByHeaderSafe: '" & headerName & "' not found"
    FindColumnByHeaderSafe = 0
    Exit Function
    
FindError:
    If DEBUG_MODE Then Debug.Print "ERROR in FindColumnByHeaderSafe: " & Err.Description & " (Number: " & Err.Number & ")"
    If DEBUG_MODE Then Debug.Print "  Looking for: '" & headerName & "', row: " & headerRow & ", column: " & i
    FindColumnByHeaderSafe = 0
End Function

' Helper function to check if sheet exists - SAFE VERSION
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Get unique building names from Units sheet - SAFE VERSION
Function GetUniqueBuildingNamesSafe(unitsWS As Worksheet, propertyCol As Long) As Collection
    On Error GoTo UniqueError
    
    Dim buildingNames As Collection
    Dim lastRow As Long
    Dim i As Long, processed As Long
    Dim buildingName As String
    
    Set buildingNames = New Collection
    lastRow = unitsWS.Cells(unitsWS.Rows.Count, propertyCol).End(xlUp).Row
    
    ' Safety check
    If lastRow > MAX_ROWS Then lastRow = MAX_ROWS
    
    If DEBUG_MODE Then Debug.Print "Processing " & lastRow & " rows for unique building names"
    
    For i = 2 To lastRow ' Start from row 2 (after headers)
        buildingName = Trim(unitsWS.Cells(i, propertyCol).Value)
        If Len(buildingName) > 0 Then
            On Error Resume Next
            buildingNames.Add buildingName, buildingName ' Key prevents duplicates
            On Error GoTo 0 ' Reset error handling
        End If
        
        processed = processed + 1
        
        ' Progress update every 500 rows
        If processed Mod 500 = 0 Then
            StatusBarMessage "Finding unique buildings... " & processed & "/" & (lastRow - 1)
            DoEvents
        End If
    Next i
    
    Set GetUniqueBuildingNamesSafe = buildingNames
    Exit Function
    
UniqueError:
    If DEBUG_MODE Then Debug.Print "Error getting unique building names: " & Err.Description
    Set GetUniqueBuildingNamesSafe = New Collection ' Return empty collection
End Function

' Create SUMIF formulas for Buildings sheet - SAFE VERSION
Sub CreateBuildingFormulasSafe(buildingsWS As Worksheet, unitsWS As Worksheet, _
                               buildingNames As Collection, propertyCol As Long)
    On Error GoTo FormulaError
    
    Dim buildingRow As Long
    Dim buildingName As Variant
    Dim buildingsProcessed As Long
    Dim bankWS As Worksheet ' Add reference to Bank Schedule for Cap Valn lookup
    Dim netAreaCol As Long, rentCol As Long, erv2023Col As Long, erv2024Col As Long
    Dim capValnValue As String
    
    Set bankWS = Worksheets("Bank Schedule")
    buildingRow = 2
    buildingsProcessed = 0
    
    If DEBUG_MODE Then Debug.Print "Creating formulas for " & buildingNames.Count & " buildings"
    
    For Each buildingName In buildingNames
        ' Progress update with detailed debugging
        StatusBarMessage "Creating formulas for building " & buildingsProcessed + 1 & "/" & buildingNames.Count & ": " & Left(buildingName, 30)
        
        If DEBUG_MODE Then Debug.Print "Processing building #" & (buildingsProcessed + 1) & ": '" & buildingName & "'"
        
        ' Building name in first column
        buildingsWS.Cells(buildingRow, 1).Value = buildingName
        
        ' Net Area formula
        If DEBUG_MODE Then Debug.Print "  Finding Net Area column..."
        netAreaCol = FindColumnByHeaderSafe(unitsWS, "Net Area", 1)
        If DEBUG_MODE Then Debug.Print "  Net Area column: " & netAreaCol
        If netAreaCol > 0 Then
            On Error Resume Next
            buildingsWS.Cells(buildingRow, 2).Formula = _
                "=SUMIF(Units!" & ColumnLetterSafe(propertyCol) & ":" & ColumnLetterSafe(propertyCol) & _
                ",""" & Replace(buildingName, """", """""") & """,Units!" & ColumnLetterSafe(netAreaCol) & ":" & ColumnLetterSafe(netAreaCol) & ")"
            If Err.Number <> 0 Then
                If DEBUG_MODE Then Debug.Print "Error with Net Area formula for " & buildingName & ": " & Err.Description
                buildingsWS.Cells(buildingRow, 2).Value = "Error: " & Err.Description
            End If
            On Error GoTo FormulaError
        Else
            buildingsWS.Cells(buildingRow, 2).Value = "Net Area column not found"
        End If
        
        ' Rent PA formula
        If DEBUG_MODE Then Debug.Print "  Finding Rent PA column..."
        rentCol = FindColumnByHeaderSafe(unitsWS, "Rent PA", 1)
        If DEBUG_MODE Then Debug.Print "  Rent PA column: " & rentCol
        If rentCol > 0 Then
            On Error Resume Next
            buildingsWS.Cells(buildingRow, 3).Formula = _
                "=SUMIF(Units!" & ColumnLetterSafe(propertyCol) & ":" & ColumnLetterSafe(propertyCol) & _
                ",""" & Replace(buildingName, """", """""") & """,Units!" & ColumnLetterSafe(rentCol) & ":" & ColumnLetterSafe(rentCol) & ")"
            If Err.Number <> 0 Then
                If DEBUG_MODE Then Debug.Print "Error with Rent PA formula for " & buildingName & ": " & Err.Description
                buildingsWS.Cells(buildingRow, 3).Value = "Error: " & Err.Description
            End If
            On Error GoTo FormulaError
        Else
            buildingsWS.Cells(buildingRow, 3).Value = "Rent PA column not found"
        End If
        
        ' 2023 ERV formula
        If DEBUG_MODE Then Debug.Print "  Finding 2023 ERV column..."
        erv2023Col = FindColumnByHeaderSafe(unitsWS, "2023 ERV", 1)
        If DEBUG_MODE Then Debug.Print "  2023 ERV column: " & erv2023Col
        If erv2023Col > 0 Then
            On Error Resume Next
            buildingsWS.Cells(buildingRow, 4).Formula = _
                "=SUMIF(Units!" & ColumnLetterSafe(propertyCol) & ":" & ColumnLetterSafe(propertyCol) & _
                ",""" & Replace(buildingName, """", """""") & """,Units!" & ColumnLetterSafe(erv2023Col) & ":" & ColumnLetterSafe(erv2023Col) & ")"
            If Err.Number <> 0 Then
                If DEBUG_MODE Then Debug.Print "Error with 2023 ERV formula for " & buildingName & ": " & Err.Description
                buildingsWS.Cells(buildingRow, 4).Value = "Error: " & Err.Description
            End If
            On Error GoTo FormulaError
        Else
            buildingsWS.Cells(buildingRow, 4).Value = "2023 ERV column not found"
        End If
        
        ' 2024 ERV formula
        If DEBUG_MODE Then Debug.Print "  Finding 2024 ERV column..."
        erv2024Col = FindColumnByHeaderSafe(unitsWS, "2024 ERV", 1)
        If DEBUG_MODE Then Debug.Print "  2024 ERV column: " & erv2024Col
        If erv2024Col > 0 Then
            On Error Resume Next
            buildingsWS.Cells(buildingRow, 5).Formula = _
                "=SUMIF(Units!" & ColumnLetterSafe(propertyCol) & ":" & ColumnLetterSafe(propertyCol) & _
                ",""" & Replace(buildingName, """", """""") & """,Units!" & ColumnLetterSafe(erv2024Col) & ":" & ColumnLetterSafe(erv2024Col) & ")"
            If Err.Number <> 0 Then
                If DEBUG_MODE Then Debug.Print "Error with 2024 ERV formula for " & buildingName & ": " & Err.Description
                buildingsWS.Cells(buildingRow, 5).Value = "Error: " & Err.Description
            End If
            On Error GoTo FormulaError
        Else
            buildingsWS.Cells(buildingRow, 5).Value = "2024 ERV column not found"
        End If
        
        ' Cap Valn with automatic detection and mapping
        If DEBUG_MODE Then Debug.Print "  Finding Cap Valn for building: " & buildingName
        capValnValue = FindCapValnForBuilding(buildingName, bankWS)
        If DEBUG_MODE Then Debug.Print "  Cap Valn result: " & capValnValue
        
        If capValnValue <> "" Then
            buildingsWS.Cells(buildingRow, 6).Formula = capValnValue
        Else
            buildingsWS.Cells(buildingRow, 6).Value = "Cap Valn not found for: " & buildingName
            buildingsWS.Cells(buildingRow, 6).Interior.Color = RGB(255, 200, 200) ' Light red for not found
        End If
        
        buildingRow = buildingRow + 1
        buildingsProcessed = buildingsProcessed + 1
        
        ' Allow Excel to respond every 10 buildings
        If buildingsProcessed Mod 10 = 0 Then DoEvents
        
        ' Safety check
        If buildingRow > MAX_ROWS Then
            MsgBox "Too many buildings! Stopping at " & MAX_ROWS & " for safety.", vbExclamation
            Exit For
        End If
    Next buildingName
    
    If DEBUG_MODE Then Debug.Print "Created formulas for " & buildingsProcessed & " buildings"
    Exit Sub
    
FormulaError:
    If DEBUG_MODE Then Debug.Print "ERROR in CreateBuildingFormulas: " & Err.Description & " (Number: " & Err.Number & ")"
    If DEBUG_MODE Then Debug.Print "  Error occurred while processing building: " & buildingName & " (building #" & (buildingsProcessed + 1) & ")"
    If DEBUG_MODE Then Debug.Print "  Building row: " & buildingRow & ", Buildings processed: " & buildingsProcessed
    MsgBox "Error creating building formulas: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Building: " & buildingName & vbCrLf & _
           "Processed " & buildingsProcessed & " buildings before error", vbCritical
End Sub

' Helper function to convert column number to letter - SAFE VERSION
Function ColumnLetterSafe(colNum As Long) As String
    On Error GoTo ColumnError
    
    If colNum <= 0 Or colNum > 16384 Then ' Excel column limit
        ColumnLetterSafe = "A" ' Default fallback
        Exit Function
    End If
    
    ColumnLetterSafe = Split(Cells(1, colNum).Address, "$")(1)
    Exit Function
    
ColumnError:
    ColumnLetterSafe = "A" ' Fallback
End Function

' Helper function for status bar messages
Sub StatusBarMessage(message As String)
    Application.StatusBar = message
    If DEBUG_MODE Then Debug.Print message
End Sub

' Find Cap Valn reference for a building - AUTOMATIC DETECTION
Function FindCapValnForBuilding(buildingName As String, bankWS As Worksheet) As String
    On Error GoTo CapValnError
    
    If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Starting for '" & buildingName & "'"
    
    ' First, find the Cap Valn column in Bank Schedule
    Dim capValnCol As Long
    If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Calling FindCapValnColumn..."
    capValnCol = FindCapValnColumn(bankWS)
    If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Cap Valn column result = " & capValnCol
    
    If capValnCol = 0 Then
        If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Cap Valn column not found"
        FindCapValnForBuilding = ""
        Exit Function
    End If
    
    If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Cap Valn column found at: " & ColumnLetterSafe(capValnCol)
    
    ' Find building rows that match this building name
    Dim matchingRows As Collection
    If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Calling FindBuildingRows..."
    Set matchingRows = FindBuildingRows(buildingName, bankWS)
    If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: Found " & matchingRows.Count & " matching rows"
    
    If matchingRows.Count = 0 Then
        If DEBUG_MODE Then Debug.Print "    FindCapValnForBuilding: No matching rows found for: " & buildingName
        FindCapValnForBuilding = ""
        Exit Function
    End If
    
    ' Look for Cap Valn value in the matching rows
    Dim rowNum As Variant
    For Each rowNum In matchingRows
        Dim capValnValue As Variant
        capValnValue = bankWS.Cells(rowNum, capValnCol).Value
        
        ' Check if this cell has a Cap Valn value (numeric and > 0)
        If IsNumeric(capValnValue) And CDbl(capValnValue) > 0 Then
            Dim cellRef As String
            cellRef = "='Bank Schedule'!" & ColumnLetterSafe(capValnCol) & rowNum
            If DEBUG_MODE Then Debug.Print "Found Cap Valn for " & buildingName & " at " & cellRef & " = " & capValnValue
            FindCapValnForBuilding = cellRef
            Exit Function
        End If
    Next rowNum
    
    ' If no direct match, try fuzzy matching
    Dim fuzzyResult As String
    fuzzyResult = FindCapValnFuzzyMatch(buildingName, bankWS, capValnCol)
    FindCapValnForBuilding = fuzzyResult
    Exit Function
    
CapValnError:
    If DEBUG_MODE Then Debug.Print "Error in FindCapValnForBuilding: " & Err.Description
    FindCapValnForBuilding = ""
End Function

' Find the Cap Valn column in Bank Schedule - SIMPLIFIED VERSION
Function FindCapValnColumn(bankWS As Worksheet) As Long
    On Error GoTo FindCapError
    
    Dim lastCol As Long, i As Long, headerRow As Long
    Dim cellValue As String
    Dim targetHeader As String
    
    ' The exact header we're looking for (without quotes)
    targetHeader = "2024 Cap Valn. (£)"
    
    If DEBUG_MODE Then Debug.Print "Looking for exact header: " & targetHeader
    
    ' Check multiple rows for the exact header
    For headerRow = 1 To 5
        lastCol = bankWS.Cells(headerRow, Columns.Count).End(xlToLeft).Column
        If lastCol > MAX_COLS Then lastCol = MAX_COLS
        
        For i = 1 To lastCol
            cellValue = Trim(bankWS.Cells(headerRow, i).Value)
            
            ' Remove extra spaces from the cell value for comparison
            cellValue = Replace(cellValue, "  ", " ")
            cellValue = Trim(cellValue)
            
            ' Exact match with the target header
            If cellValue = targetHeader Then
                If DEBUG_MODE Then Debug.Print "Found exact match at row " & headerRow & ", column " & i & " (" & ColumnLetterSafe(i) & ")"
                FindCapValnColumn = i
                Exit Function
            End If
        Next i
    Next headerRow
    
    If DEBUG_MODE Then Debug.Print "Exact header '" & targetHeader & "' not found"
    FindCapValnColumn = 0
    Exit Function
    
FindCapError:
    If DEBUG_MODE Then Debug.Print "Error in FindCapValnColumn: " & Err.Description
    FindCapValnColumn = 0
End Function

' Find rows in Bank Schedule that contain building names (not unit data)
Function FindBuildingRows(buildingName As String, bankWS As Worksheet) As Collection
    On Error GoTo BuildingRowsError
    
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Starting for '" & buildingName & "'"
    
    Dim matchingRows As Collection
    Set matchingRows = New Collection
    
    Dim lastRow As Long, i As Long, checkCol As Long
    Dim unitDemiseCol As Long, propertyCol As Long
    
    ' Find the relevant columns
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Finding Unit Demise column..."
    unitDemiseCol = FindColumnByHeaderSafe(bankWS, "Unit Demise", 3)
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Unit Demise column = " & unitDemiseCol
    
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Finding Property column..."
    propertyCol = FindColumnByHeaderSafe(bankWS, "Property", 3)
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Property column = " & propertyCol
    
    ' Get data bounds
    lastRow = 0
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Getting data bounds..."
    For checkCol = 1 To 10
        Dim colLastRow As Long
        colLastRow = bankWS.Cells(bankWS.Rows.Count, checkCol).End(xlUp).Row
        If colLastRow > lastRow Then lastRow = colLastRow
    Next checkCol
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Last row = " & lastRow
    
    ' Look through all rows for building matches
    If DEBUG_MODE Then Debug.Print "      FindBuildingRows: Starting row scan from 4 to " & lastRow
    For i = 4 To lastRow ' Start after headers
        If DEBUG_MODE And (i Mod 100 = 0) Then Debug.Print "      FindBuildingRows: Processing row " & i
        ' Skip unit rows (those with Unit Demise data)
        Dim isUnitRow As Boolean
        isUnitRow = False
        If unitDemiseCol > 0 Then
            isUnitRow = (Len(Trim(bankWS.Cells(i, unitDemiseCol).Value)) > 0)
        End If
        
        If Not isUnitRow Then
            ' Check multiple columns for building name matches
            Dim found As Boolean
            found = False
            
            ' Check Property column
            If propertyCol > 0 Then
                If IsBuildingNameMatch(buildingName, Trim(bankWS.Cells(i, propertyCol).Value)) Then
                    found = True
                End If
            End If
            
            ' Check first few columns for building names
            If Not found Then
                Dim maxCheckCol As Long
                maxCheckCol = bankWS.Cells(3, Columns.Count).End(xlToLeft).Column
                If maxCheckCol > 10 Then maxCheckCol = 10
                
                For checkCol = 1 To maxCheckCol
                    Dim cellValue As String
                    cellValue = Trim(bankWS.Cells(i, checkCol).Value)
                    If IsBuildingNameMatch(buildingName, cellValue) Then
                        found = True
                        Exit For
                    End If
                Next checkCol
            End If
            
            If found Then
                On Error Resume Next
                matchingRows.Add i, CStr(i)
                On Error GoTo BuildingRowsError
            End If
        End If
    Next i
    
    Set FindBuildingRows = matchingRows
    Exit Function
    
BuildingRowsError:
    If DEBUG_MODE Then Debug.Print "ERROR in FindBuildingRows: " & Err.Description & " (Number: " & Err.Number & ") at row " & i
    Set FindBuildingRows = New Collection
End Function

' Check if two building names match (with fuzzy logic)
Function IsBuildingNameMatch(name1 As String, name2 As String) As Boolean
    On Error Resume Next
    
    If Len(name1) = 0 Or Len(name2) = 0 Then
        IsBuildingNameMatch = False
        Exit Function
    End If
    
    Dim clean1 As String, clean2 As String
    
    ' Clean the names for comparison
    clean1 = CleanBuildingName(name1)
    clean2 = CleanBuildingName(name2)
    
    ' Exact match after cleaning
    If clean1 = clean2 Then
        IsBuildingNameMatch = True
        Exit Function
    End If
    
    ' Contains match (either direction)
    If InStr(1, clean1, clean2, vbTextCompare) > 0 Or InStr(1, clean2, clean1, vbTextCompare) > 0 Then
        IsBuildingNameMatch = True
        Exit Function
    End If
    
    ' Check for common variations
    If CheckBuildingVariations(clean1, clean2) Then
        IsBuildingNameMatch = True
        Exit Function
    End If
    
    IsBuildingNameMatch = False
End Function

' Clean building names for comparison
Function CleanBuildingName(buildingName As String) As String
    Dim result As String
    result = UCase(Trim(buildingName))
    
    ' Remove common words
    result = Replace(result, "BUILDING", "BLDG")
    result = Replace(result, "BLOCK", "BLK")
    result = Replace(result, "TOWER", "TWR")
    result = Replace(result, "CENTRE", "CENTER")
    result = Replace(result, "STREET", "ST")
    result = Replace(result, "ROAD", "RD")
    result = Replace(result, "AVENUE", "AVE")
    
    ' Remove extra spaces and punctuation
    result = Replace(result, "  ", " ")
    result = Replace(result, ".", "")
    result = Replace(result, ",", "")
    result = Replace(result, "-", " ")
    result = Replace(result, "_", " ")
    result = Trim(result)
    
    CleanBuildingName = result
End Function

' Check for common building name variations
Function CheckBuildingVariations(name1 As String, name2 As String) As Boolean
    ' Check if one is a subset of the other with different formatting
    Dim words1() As String, words2() As String
    Dim i As Long, j As Long, matches As Long
    
    words1 = Split(name1, " ")
    words2 = Split(name2, " ")
    
    matches = 0
    
    ' Count matching words
    For i = 0 To UBound(words1)
        For j = 0 To UBound(words2)
            If Len(words1(i)) > 2 And words1(i) = words2(j) Then
                matches = matches + 1
                Exit For
            End If
        Next j
    Next i
    
    ' Consider it a match if most significant words match
    Dim minWords As Long
    Dim bound1 As Long, bound2 As Long
    bound1 = UBound(words1) + 1
    bound2 = UBound(words2) + 1
    If bound1 < bound2 Then
        minWords = bound1
    Else
        minWords = bound2
    End If
    
    CheckBuildingVariations = (matches >= minWords * 0.6) ' 60% word match threshold
End Function

' Fuzzy match for Cap Valn when exact match fails
Function FindCapValnFuzzyMatch(buildingName As String, bankWS As Worksheet, capValnCol As Long) As String
    On Error GoTo FuzzyError
    
    If DEBUG_MODE Then Debug.Print "Trying fuzzy match for: " & buildingName
    
    Dim bestMatch As String
    Dim bestScore As Double
    Dim bestRow As Long
    
    bestScore = 0
    bestMatch = ""
    
    ' Search through all rows for potential matches
    Dim lastRow As Long, i As Long
    lastRow = bankWS.Cells(bankWS.Rows.Count, 1).End(xlUp).Row
    
    For i = 4 To lastRow
        ' Skip if this row has a Unit Demise (it's a unit row)
        Dim unitDemiseCol As Long
        unitDemiseCol = FindColumnByHeaderSafe(bankWS, "Unit Demise", 3)
        If unitDemiseCol > 0 And Len(Trim(bankWS.Cells(i, unitDemiseCol).Value)) > 0 Then
            GoTo NextRow
        End If
        
        ' Check if this row has a Cap Valn value
        Dim capValnValue As Variant
        capValnValue = bankWS.Cells(i, capValnCol).Value
        If Not IsNumeric(capValnValue) Or CDbl(capValnValue) <= 0 Then
            GoTo NextRow
        End If
        
        ' Check all text cells in this row for building name matches
        Dim lastCol As Long, j As Long
        lastCol = bankWS.Cells(i, Columns.Count).End(xlToLeft).Column
        If lastCol > 20 Then lastCol = 20 ' Limit search range
        
        For j = 1 To lastCol
            Dim cellValue As String
            cellValue = Trim(bankWS.Cells(i, j).Value)
            
            If Len(cellValue) > 0 Then
                Dim similarity As Double
                similarity = CalculateSimilarity(buildingName, cellValue)
                
                If similarity > bestScore And similarity > 0.7 Then ' 70% similarity threshold
                    bestScore = similarity
                    bestMatch = "='Bank Schedule'!" & ColumnLetterSafe(capValnCol) & i
                    bestRow = i
                End If
            End If
        Next j
        
NextRow:
    Next i
    
    If bestScore > 0.7 Then
        If DEBUG_MODE Then Debug.Print "Fuzzy match found for " & buildingName & " at row " & bestRow & " with score " & Format(bestScore, "0.00")
        FindCapValnFuzzyMatch = bestMatch
    Else
        If DEBUG_MODE Then Debug.Print "No fuzzy match found for " & buildingName
        FindCapValnFuzzyMatch = ""
    End If
    
    Exit Function
    
FuzzyError:
    FindCapValnFuzzyMatch = ""
End Function

' Calculate similarity between two strings (simple implementation)
Function CalculateSimilarity(str1 As String, str2 As String) As Double
    On Error Resume Next
    
    If Len(str1) = 0 Or Len(str2) = 0 Then
        CalculateSimilarity = 0
        Exit Function
    End If
    
    Dim clean1 As String, clean2 As String
    clean1 = CleanBuildingName(str1)
    clean2 = CleanBuildingName(str2)
    
    ' Exact match
    If clean1 = clean2 Then
        CalculateSimilarity = 1
        Exit Function
    End If
    
    ' Contains match
    If InStr(1, clean1, clean2, vbTextCompare) > 0 Or InStr(1, clean2, clean1, vbTextCompare) > 0 Then
        CalculateSimilarity = 0.8
        Exit Function
    End If
    
    ' Word-based similarity
    Dim words1() As String, words2() As String
    Dim commonWords As Long, totalWords As Long
    
    words1 = Split(clean1, " ")
    words2 = Split(clean2, " ")
    
    commonWords = 0
    totalWords = UBound(words1) + 1 + UBound(words2) + 1
    
    Dim i As Long, j As Long
    For i = 0 To UBound(words1)
        For j = 0 To UBound(words2)
            If Len(words1(i)) > 2 And words1(i) = words2(j) Then
                commonWords = commonWords + 2 ' Count twice (once for each string)
                Exit For
            End If
        Next j
    Next i
    
    CalculateSimilarity = commonWords / totalWords
End Function

' Helper function to get minimum value
Function Min(val1 As Long, val2 As Long) As Long
    If val1 < val2 Then
        Min = val1
    Else
        Min = val2
    End If
End Function

' Auto-run on workbook open (optional) - SAFE VERSION
Private Sub Workbook_Open()
    On Error Resume Next ' Don't let auto-run crash the workbook
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Process Bank Schedule data into Units and Buildings sheets?" & vbCrLf & _
                      "Note: This will take a few moments and show progress messages.", _
                      vbYesNo + vbQuestion, "Bank Schedule Sanitizer")
    
    If response = vbYes Then
        Call ProcessBankSchedule
    End If
End Sub
