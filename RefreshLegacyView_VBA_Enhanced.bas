Sub RefreshLegacyView()
    ' Dynamic Legacy View Rebuilder - Enhanced with Error Handling
    ' Parses Units sheet and recreates hierarchical Legacy View structure
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Dim unitsWs As Worksheet
    Dim legacyWs As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim headerRow As Long
    Dim col As Long
    Dim buildingName As String
    Dim i As Long
    Dim j As Long
    Dim startRow As Long, endRow As Long
    Dim wsExists As Boolean
    Dim buildingIndex As Long
    Dim unitIndex As Long
    
    ' Check if Units sheet exists
    wsExists = False
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name = "Units" Then
            wsExists = True
            Exit For
        End If
    Next i
    
    If Not wsExists Then
        MsgBox "Error: 'Units' sheet not found in this workbook." & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical, "Sheet Not Found"
        GoTo Cleanup
    End If
    
    ' Get Units worksheet
    Set unitsWs = ThisWorkbook.Worksheets("Units")
    
    ' Check if Legacy View sheet exists, create if it doesn't
    wsExists = False
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name = "Legacy View" Then
            wsExists = True
            Exit For
        End If
    Next i
    
    If Not wsExists Then
        ' Create Legacy View sheet
        Set legacyWs = ThisWorkbook.Worksheets.Add
        legacyWs.Name = "Legacy View"
        ' Add title and instructions
        legacyWs.Cells(1, 1).Value = "Legacy View - Dynamic Hierarchical Structure"
        legacyWs.Cells(2, 1).Value = "(Auto-rebuilds from Units sheet - responsive to structure changes)"
        legacyWs.Cells(3, 1).Value = "Press Ctrl+Shift+R to refresh or run RefreshLegacyView macro"
    Else
        Set legacyWs = ThisWorkbook.Worksheets("Legacy View")
    End If
    
    ' Clear existing Legacy View content (keep first 3 instruction rows)
    lastRow = legacyWs.Cells(legacyWs.Rows.Count, 1).End(xlUp).Row
    If lastRow > 3 Then
        legacyWs.Range("A4:ZZ" & lastRow).Clear
    End If
    
    ' Check if Units sheet has data
    If unitsWs.Cells(1, 1).Value = "" Then
        MsgBox "Error: Units sheet appears to be empty or has no headers.", vbCritical
        GoTo Cleanup
    End If
    
    ' Set up headers
    headerRow = 4
    Dim maxCol As Long
    maxCol = unitsWs.Cells(1, unitsWs.Columns.Count).End(xlToLeft).Column
    
    ' Copy headers from Units sheet (skip Building column which is column 1)
    Dim targetCol As Long
    targetCol = 1
    For col = 2 To maxCol
        If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
            legacyWs.Cells(headerRow, targetCol).Value = unitsWs.Cells(1, col).Value
            legacyWs.Cells(headerRow, targetCol).Font.Bold = True
            targetCol = targetCol + 1
        End If
    Next col
    
    ' Group units by building using arrays instead of Dictionary
    lastRow = unitsWs.Cells(unitsWs.Rows.Count, 1).End(xlUp).Row
    
    ' Check if there's actual data beyond headers
    If lastRow < 2 Then
        MsgBox "Error: No data found in Units sheet beyond headers.", vbCritical
        GoTo Cleanup
    End If
    
    ' First pass: collect unique building names
    Dim buildingNames() As String
    Dim buildingCount As Long
    buildingCount = 0
    
    For i = 2 To lastRow
        buildingName = Trim(CStr(unitsWs.Cells(i, 1).Value))
        If buildingName <> "" And buildingName <> "0" Then
            ' Check if building already exists
            Dim found As Boolean
            found = False
            For j = 1 To buildingCount
                If buildingNames(j) = buildingName Then
                    found = True
                    Exit For
                End If
            Next j
            
            ' Add new building if not found
            If Not found Then
                buildingCount = buildingCount + 1
                ReDim Preserve buildingNames(1 To buildingCount)
                buildingNames(buildingCount) = buildingName
            End If
        End If
    Next i
    
    ' Check if we found any buildings
    If buildingCount = 0 Then
        MsgBox "Error: No valid building names found in Units sheet column A.", vbCritical
        GoTo Cleanup
    End If
    
    ' Build hierarchical structure
    currentRow = headerRow + 2 ' Empty row after headers
    
    ' Process each building
    For buildingIndex = 1 To buildingCount
        Dim currentBuilding As String
        currentBuilding = buildingNames(buildingIndex)
        
        ' Collect unit rows for this building
        Dim unitRowsForBuilding() As Long
        Dim unitRowCount As Long
        unitRowCount = 0
        
        For i = 2 To lastRow
            If Trim(CStr(unitsWs.Cells(i, 1).Value)) = currentBuilding Then
                unitRowCount = unitRowCount + 1
                ReDim Preserve unitRowsForBuilding(1 To unitRowCount)
                unitRowsForBuilding(unitRowCount) = i
            End If
        Next i
        
        ' Building header
        legacyWs.Cells(currentRow, 1).Value = currentBuilding
        legacyWs.Cells(currentRow, 1).Font.Bold = True
        legacyWs.Cells(currentRow, 1).Font.Size = 12
        currentRow = currentRow + 2 ' Empty row after building header
        
        ' Unit rows
        startRow = currentRow
        For unitIndex = 1 To unitRowCount
            targetCol = 1
            For col = 2 To maxCol
                If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
                    legacyWs.Cells(currentRow, targetCol).Formula = "=Units!" & _
                        ConvertToColumnLetter(col) & unitRowsForBuilding(unitIndex)
                    targetCol = targetCol + 1
                End If
            Next col
            currentRow = currentRow + 1
        Next unitIndex
        endRow = currentRow - 1
        
        ' Empty row + summary row
        currentRow = currentRow + 1
        legacyWs.Cells(currentRow, 1).Value = currentBuilding & " - TOTAL"
        legacyWs.Cells(currentRow, 1).Font.Bold = True
        legacyWs.Cells(currentRow, 1).Font.Italic = True
        
        ' Add SUM formulas for numeric columns
        targetCol = 1
        For col = 2 To maxCol
            If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
                If IsNumericColumn(unitsWs.Cells(1, col).Value) Then
                    If startRow <= endRow Then
                        legacyWs.Cells(currentRow, targetCol).Formula = "=SUM(" & _
                            ConvertToColumnLetter(targetCol) & startRow & ":" & _
                            ConvertToColumnLetter(targetCol) & endRow & ")"
                        legacyWs.Cells(currentRow, targetCol).Font.Bold = True
                    End If
                End If
                targetCol = targetCol + 1
            End If
        Next col
        
        currentRow = currentRow + 2 ' Empty row after summary
    Next buildingIndex
    
    ' Auto-fit columns
    legacyWs.Columns.AutoFit
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Legacy View refreshed successfully!" & vbCrLf & _
           "Found " & buildingCount & " buildings with " & (lastRow - 1) & " total units.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing Legacy View: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Available sheets: " & GetSheetNames(), vbCritical
    GoTo Cleanup
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function ConvertToColumnLetter(colNum As Long) As String
    Dim result As String
    result = ""
    While colNum > 0
        colNum = colNum - 1
        result = Chr(65 + (colNum Mod 26)) & result
        colNum = colNum \ 26
    Wend
    ConvertToColumnLetter = result
End Function

Private Function IsNumericColumn(header As String) As Boolean
    Dim lowerHeader As String
    lowerHeader = LCase(Trim(header))
    
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
                      (InStr(lowerHeader, "cost") > 0)
End Function

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

Private Sub Auto_Open()
    ' Set up keyboard shortcut when workbook opens
    Application.OnKey "^+R", "RefreshLegacyView"
End Sub

Sub TestSheetExists()
    ' Helper macro to test what sheets exist
    MsgBox "Available sheets in this workbook:" & vbCrLf & vbCrLf & GetSheetNames(), vbInformation
End Sub
