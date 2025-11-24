Sub RefreshLegacyView()
    ' Dynamic Legacy View Rebuilder - Full Hierarchical Structure with Efficient Processing
    ' Creates complete building structure with all units and data
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    MsgBox "HIERARCHICAL: Starting full structure creation...", vbInformation
    
    Dim unitsWs As Worksheet
    Dim legacyWs As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim headerRow As Long
    Dim col As Long
    Dim buildingName As String
    Dim i As Long
    Dim j As Long
    Dim buildingIndex As Long
    Dim unitIndex As Long
    Dim startRow As Long, endRow As Long
    
    ' Get Units worksheet
    Set unitsWs = ThisWorkbook.Worksheets("Units")
    
    ' Remove and recreate Legacy View for clean start
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Legacy View").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set legacyWs = ThisWorkbook.Worksheets.Add
    legacyWs.Name = "Legacy View"
    
    MsgBox "HIERARCHICAL: Created fresh Legacy View sheet", vbInformation
    
    ' Get data dimensions
    lastRow = unitsWs.Cells(unitsWs.Rows.Count, 1).End(xlUp).Row
    Dim maxCol As Long
    maxCol = unitsWs.Cells(1, unitsWs.Columns.Count).End(xlToLeft).Column
    
    MsgBox "HIERARCHICAL: Processing " & lastRow & " rows x " & maxCol & " columns", vbInformation
    
    ' Set up headers to match Bank Schedule structure - start at row 3
    headerRow = 3
    Dim targetCol As Long
    
    ' Column 1: Empty (to match Bank Schedule)
    legacyWs.Cells(headerRow, 1).Value = ""
    
    ' Start mapping from column 2 - no gap
    targetCol = 2
    For col = 2 To maxCol
        If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
            legacyWs.Cells(headerRow, targetCol).Value = unitsWs.Cells(1, col).Value
            On Error Resume Next
            legacyWs.Cells(headerRow, targetCol).Font.Bold = True
            legacyWs.Cells(headerRow, targetCol).Font.Size = unitsWs.Cells(1, col).Font.Size
            legacyWs.Cells(headerRow, targetCol).Font.Name = unitsWs.Cells(1, col).Font.Name
            legacyWs.Cells(headerRow, targetCol).Interior.Color = unitsWs.Cells(1, col).Interior.Color
            legacyWs.Cells(headerRow, targetCol).NumberFormat = unitsWs.Cells(1, col).NumberFormat
            Err.Clear
            On Error GoTo ErrorHandler
            targetCol = targetCol + 1
        End If
    Next col
    
    Dim numDataCols As Long
    numDataCols = targetCol - 1
    
    ' Hide columns 3-5 (Client, Property Number, Property)
    On Error Resume Next
    legacyWs.Columns("C:E").Hidden = True
    Err.Clear
    On Error GoTo ErrorHandler
    
    MsgBox "HIERARCHICAL: Headers set up. " & numDataCols & " data columns", vbInformation
    
    ' Collect unique buildings efficiently
    Dim buildingNames() As String
    Dim buildingCount As Long
    buildingCount = 0
    
    For i = 2 To lastRow
        buildingName = Trim(CStr(unitsWs.Cells(i, 1).Value))
        If buildingName <> "" And buildingName <> "0" And buildingName <> "False" Then
            Dim buildingExists As Boolean
            buildingExists = False
            For j = 1 To buildingCount
                If buildingNames(j) = buildingName Then
                    buildingExists = True
                    Exit For
                End If
            Next j
            
            If Not buildingExists Then
                buildingCount = buildingCount + 1
                ReDim Preserve buildingNames(1 To buildingCount)
                buildingNames(buildingCount) = buildingName
            End If
        End If
    Next i
    
    MsgBox "HIERARCHICAL: Found " & buildingCount & " buildings. Creating structure...", vbInformation
    
    ' Create hierarchical structure - process in smaller batches
    currentRow = headerRow + 2 ' Start after headers with empty row
    
    For buildingIndex = 1 To buildingCount
        Dim currentBuilding As String
        currentBuilding = buildingNames(buildingIndex)
        
        ' Show progress for every 10 buildings
        If buildingIndex Mod 10 = 1 Or buildingIndex = 1 Then
            MsgBox "HIERARCHICAL: Processing building " & buildingIndex & " of " & buildingCount & vbCrLf & _
                   "Current: " & currentBuilding, vbInformation
        End If
        
        ' Collect unit rows for this building
        Dim unitRows() As Long
        Dim unitCount As Long
        unitCount = 0
        
        For i = 2 To lastRow
            If Trim(CStr(unitsWs.Cells(i, 1).Value)) = currentBuilding Then
                unitCount = unitCount + 1
                ReDim Preserve unitRows(1 To unitCount)
                unitRows(unitCount) = i
            End If
        Next i
        
        ' Only process buildings that have units
        If unitCount > 0 Then
            ' Building header row - put building name in Unit Type column (column 7)
            ' With our simplified mapping: Surveyor(B=2), Client(C=3), PropNum(D=4), Prop(E=5), UnitDemise(F=6), UnitType(G=7)
            On Error Resume Next
            legacyWs.Cells(currentRow, 7).Value = currentBuilding
            legacyWs.Cells(currentRow, 7).Font.Bold = True
            legacyWs.Cells(currentRow, 7).Font.Size = 12
            Err.Clear
            On Error GoTo ErrorHandler
            currentRow = currentRow + 2 ' Empty row after building name
            
            ' Process units using efficient batch copying instead of individual formulas
            startRow = currentRow
            
            ' Copy unit data in batches to improve performance
            For unitIndex = 1 To unitCount
                Dim sourceRow As Long
                sourceRow = unitRows(unitIndex)
                
                ' Copy values, formatting, and data types with correct column mapping
                ' Column 1: Empty (matches Bank Schedule)
                ' Copy all columns from source to target with simple mapping
                targetCol = 2 ' Start at column 2
                For col = 2 To maxCol
                    If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
                        On Error Resume Next
                        sourceValue = unitsWs.Cells(sourceRow, col).Value
                        
                        If Err.Number = 0 Then
                            legacyWs.Cells(currentRow, targetCol).Value = sourceValue
                            legacyWs.Cells(currentRow, targetCol).NumberFormat = unitsWs.Cells(sourceRow, col).NumberFormat
                            legacyWs.Cells(currentRow, targetCol).Font.Bold = unitsWs.Cells(sourceRow, col).Font.Bold
                            legacyWs.Cells(currentRow, targetCol).Font.Italic = unitsWs.Cells(sourceRow, col).Font.Italic
                            legacyWs.Cells(currentRow, targetCol).Font.Size = unitsWs.Cells(sourceRow, col).Font.Size
                            legacyWs.Cells(currentRow, targetCol).Font.Name = unitsWs.Cells(sourceRow, col).Font.Name
                            legacyWs.Cells(currentRow, targetCol).Font.Color = unitsWs.Cells(sourceRow, col).Font.Color
                            legacyWs.Cells(currentRow, targetCol).Interior.Color = unitsWs.Cells(sourceRow, col).Interior.Color
                            legacyWs.Cells(currentRow, targetCol).HorizontalAlignment = unitsWs.Cells(sourceRow, col).HorizontalAlignment
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                        targetCol = targetCol + 1
                    End If
                Next col
                currentRow = currentRow + 1
                
                ' Process in small batches to prevent freezing
                If unitIndex Mod 10 = 0 Then
                    DoEvents ' Allow Excel to respond
                End If
            Next unitIndex
            
            endRow = currentRow - 1
            
            ' Add empty row before summary
            currentRow = currentRow + 1
            
            ' Building summary row with SUM formulas
            On Error Resume Next
            legacyWs.Cells(currentRow, 7).Value = currentBuilding & " - TOTAL"
            legacyWs.Cells(currentRow, 7).Font.Bold = True
            legacyWs.Cells(currentRow, 7).Font.Italic = True
            Err.Clear
            On Error GoTo ErrorHandler
            
            ' Add SUM formulas for ALL numeric columns in summary row
            ' Reset targetCol to match the column mapping used when copying data
            targetCol = 2 ' Start at column 2 to match our simplified mapping
            For col = 2 To maxCol
                On Error Resume Next
                Dim header As Variant
                header = unitsWs.Cells(1, col).Value
                On Error GoTo ErrorHandler
                
                If Not IsEmpty(header) And header <> "" Then
                    ' Check if this is a numeric column that should be summed
                    If IsNumericColumn(CStr(header)) And startRow <= endRow And startRow < endRow Then
                        Dim colLetter As String
                        colLetter = ConvertToColumnLetter(targetCol)
                        
                        ' Only create SUM if we have multiple rows to sum and targetCol is valid
                        If startRow < endRow And targetCol > 0 Then
                            On Error Resume Next
                            legacyWs.Cells(currentRow, targetCol).Formula = "=SUM(" & _
                                colLetter & startRow & ":" & colLetter & endRow & ")"
                            If Err.Number = 0 Then
                                legacyWs.Cells(currentRow, targetCol).Font.Bold = True
                                ' Copy number formatting from the first unit row for consistency
                                If startRow > 0 And startRow <= endRow Then
                                    legacyWs.Cells(currentRow, targetCol).NumberFormat = _
                                        legacyWs.Cells(startRow, targetCol).NumberFormat
                                End If
                                ' Only show success message for first building to confirm it's working
                                If buildingIndex = 1 Then
                                    MsgBox "DEBUG: Successfully added SUM formula for " & header & " in " & colLetter & currentRow, vbInformation
                                End If
                            Else
                                ' Always show errors
                                MsgBox "ERROR: Failed to add SUM formula for " & header & ": " & Err.Description, vbCritical
                            End If
                            Err.Clear
                            On Error GoTo ErrorHandler
                        End If
                    End If
                    ' CRITICAL: Always increment targetCol for every non-empty header to maintain mapping consistency
                    targetCol = targetCol + 1
                End If
            Next col
            
            currentRow = currentRow + 2 ' Empty row after summary
        End If
        
        ' Allow Excel to process and show progress
        If buildingIndex Mod 5 = 0 Then
            DoEvents
        End If
    Next buildingIndex
    
    MsgBox "HIERARCHICAL: Structure complete. Formatting and finalizing...", vbInformation
    
    ' Copy column widths from Units sheet for better visual consistency
    On Error Resume Next
    legacyWs.Columns(1).ColumnWidth = 2 ' Empty column - narrow
    
    ' Copy widths for all other columns using simple mapping
    targetCol = 2
    For col = 2 To maxCol
        If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
            legacyWs.Columns(targetCol).ColumnWidth = unitsWs.Columns(col).ColumnWidth
            targetCol = targetCol + 1
        End If
    Next col
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Auto-fit columns if copying widths didn't work well
    On Error Resume Next
    legacyWs.Columns.AutoFit
    On Error GoTo ErrorHandler
    
    ' Apply some basic formatting to make it readable
    On Error Resume Next
    legacyWs.Range("A" & headerRow & ":" & ConvertToColumnLetter(numDataCols) & headerRow).Font.Bold = True
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Add consistent light borders to all data
    On Error Resume Next
    Dim lastDataRow As Long
    lastDataRow = legacyWs.Cells(legacyWs.Rows.Count, 1).End(xlUp).Row
    
    With legacyWs.Range("A" & headerRow & ":" & ConvertToColumnLetter(numDataCols) & lastDataRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200) ' Light gray borders
    End With
    Err.Clear
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "HIERARCHICAL: Legacy View created successfully!" & vbCrLf & _
           "Full hierarchical structure with " & buildingCount & " buildings" & vbCrLf & _
           "Contains all original data with building summaries" & vbCrLf & vbCrLf & _
           "Note: Uses static values for performance. Run macro again to refresh.", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description & " (Number: " & Err.Number & ")", vbCritical
End Sub

Private Function IsNumericColumn(header As String) As Boolean
    If IsEmpty(header) Then
        IsNumericColumn = False
        Exit Function
    End If
    
    Dim lowerHeader As String
    lowerHeader = LCase(Trim(CStr(header)))
    
    ' Check for common numeric column patterns
    IsNumericColumn = (InStr(lowerHeader, "area") > 0) Or _
                      (InStr(lowerHeader, "rent") > 0) Or _
                      (InStr(lowerHeader, "erv") > 0) Or _
                      (InStr(lowerHeader, "value") > 0) Or _
                      (InStr(lowerHeader, "valuation") > 0) Or _
                      (InStr(lowerHeader, "cap") > 0) Or _
                      (InStr(lowerHeader, "£") > 0) Or _
                      (InStr(lowerHeader, "$") > 0) Or _
                      (InStr(lowerHeader, "sq") > 0) Or _
                      (InStr(lowerHeader, "total") > 0) Or _
                      (InStr(lowerHeader, "sum") > 0) Or _
                      (InStr(lowerHeader, "cost") > 0) Or _
                      (InStr(lowerHeader, "price") > 0) Or _
                      (InStr(lowerHeader, "amount") > 0) Or _
                      (InStr(lowerHeader, "net") > 0) Or _
                      (InStr(lowerHeader, "gross") > 0) Or _
                      (InStr(lowerHeader, "income") > 0) Or _
                      (InStr(lowerHeader, "yield") > 0) Or _
                      (InStr(lowerHeader, "rate") > 0)
End Function

Private Function ConvertToColumnLetter(colNum As Long) As String
    If colNum <= 0 Then
        ConvertToColumnLetter = "ERROR"
        Exit Function
    End If
    
    Dim result As String
    result = ""
    While colNum > 0
        colNum = colNum - 1
        result = Chr(65 + (colNum Mod 26)) & result
        colNum = colNum \ 26
    Wend
    ConvertToColumnLetter = result
End Function

Sub TestSheetExists()
    MsgBox "Available sheets: " & GetSheetNames(), vbInformation
End Sub

Private Function GetSheetNames() As String
    Dim i As Integer
    Dim sheetList As String
    sheetList = ""
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        If i > 1 Then sheetList = sheetList & ", "
        sheetList = sheetList & ThisWorkbook.Worksheets(i).Name
    Next i
    
    GetSheetNames = sheetList
End Function
