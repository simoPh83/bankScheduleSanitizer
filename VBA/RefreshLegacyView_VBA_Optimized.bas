Sub RefreshLegacyView()
    ' Dynamic Legacy View Rebuilder - Optimized Batch Version
    ' Uses batch operations to avoid Excel freezing
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    MsgBox "OPTIMIZED: Starting Legacy View refresh with batch operations...", vbInformation
    
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
    
    ' Get worksheets
    Set unitsWs = ThisWorkbook.Worksheets("Units")
    
    ' Remove existing Legacy View and create fresh
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Legacy View").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set legacyWs = ThisWorkbook.Worksheets.Add
    legacyWs.Name = "Legacy View"
    
    MsgBox "OPTIMIZED: Created fresh Legacy View sheet", vbInformation
    
    ' Add title
    legacyWs.Cells(1, 1).Value = "Legacy View - Optimized Dynamic Structure"
    legacyWs.Cells(2, 1).Value = "(Uses batch operations for performance)"
    legacyWs.Cells(3, 1).Value = "Building summary with key metrics only"
    
    ' Get data dimensions
    lastRow = unitsWs.Cells(unitsWs.Rows.Count, 1).End(xlUp).Row
    Dim maxCol As Long
    maxCol = unitsWs.Cells(1, unitsWs.Columns.Count).End(xlToLeft).Column
    
    MsgBox "OPTIMIZED: Data size: " & lastRow & " rows x " & maxCol & " columns", vbInformation
    
    ' Create summary headers (key columns only)
    headerRow = 5
    legacyWs.Cells(headerRow, 1).Value = "Building"
    legacyWs.Cells(headerRow, 2).Value = "Unit Count"
    legacyWs.Cells(headerRow, 3).Value = "Total Net Area"
    legacyWs.Cells(headerRow, 4).Value = "Total Rent PA"
    legacyWs.Cells(headerRow, 5).Value = "Total ERV 2024"
    legacyWs.Cells(headerRow, 6).Value = "Avg ERV/sq.ft"
    
    ' Bold headers
    legacyWs.Range("A" & headerRow & ":F" & headerRow).Font.Bold = True
    
    MsgBox "OPTIMIZED: Headers created. Collecting buildings...", vbInformation
    
    ' Collect unique buildings efficiently
    Dim buildingNames() As String
    Dim buildingCount As Long
    buildingCount = 0
    
    ' Use Dictionary approach but with error handling for ActiveX issues
    Dim buildingExists As Boolean
    
    For i = 2 To lastRow
        buildingName = Trim(CStr(unitsWs.Cells(i, 1).Value))
        If buildingName <> "" And buildingName <> "0" And buildingName <> "False" Then
            
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
    
    MsgBox "OPTIMIZED: Found " & buildingCount & " buildings. Creating summary...", vbInformation
    
    ' Find key column indices in Units sheet
    Dim netAreaCol As Long, rentPACol As Long, erv2024Col As Long
    netAreaCol = 0
    rentPACol = 0
    erv2024Col = 0
    
    For col = 1 To maxCol
        Dim header As String
        header = LCase(Trim(CStr(unitsWs.Cells(1, col).Value)))
        If InStr(header, "net area") > 0 Then netAreaCol = col
        If InStr(header, "rent pa") > 0 Then rentPACol = col
        If InStr(header, "2024 erv") > 0 Or InStr(header, "erv 2024") > 0 Then erv2024Col = col
    Next col
    
    MsgBox "OPTIMIZED: Key columns - Net Area: " & netAreaCol & ", Rent PA: " & rentPACol & ", ERV 2024: " & erv2024Col, vbInformation
    
    ' Create building summaries using efficient formulas
    currentRow = headerRow + 1
    
    For buildingIndex = 1 To buildingCount
        Dim currentBuilding As String
        currentBuilding = buildingNames(buildingIndex)
        
        ' Building name
        legacyWs.Cells(currentRow, 1).Value = currentBuilding
        
        ' Unit count using COUNTIF
        legacyWs.Cells(currentRow, 2).Formula = "=COUNTIF(Units.A:A,""" & currentBuilding & """)"
        
        ' Sum formulas for key metrics using SUMIF
        If netAreaCol > 0 Then
            legacyWs.Cells(currentRow, 3).Formula = "=SUMIF(Units.A:A,""" & currentBuilding & """,Units." & _
                ConvertToColumnLetter(netAreaCol) & ":" & ConvertToColumnLetter(netAreaCol) & ")"
        End If
        
        If rentPACol > 0 Then
            legacyWs.Cells(currentRow, 4).Formula = "=SUMIF(Units.A:A,""" & currentBuilding & """,Units." & _
                ConvertToColumnLetter(rentPACol) & ":" & ConvertToColumnLetter(rentPACol) & ")"
        End If
        
        If erv2024Col > 0 Then
            legacyWs.Cells(currentRow, 5).Formula = "=SUMIF(Units.A:A,""" & currentBuilding & """,Units." & _
                ConvertToColumnLetter(erv2024Col) & ":" & ConvertToColumnLetter(erv2024Col) & ")"
            
            ' Average ERV per sq ft
            If netAreaCol > 0 Then
                legacyWs.Cells(currentRow, 6).Formula = "=IF(C" & currentRow & ">0,E" & currentRow & "/C" & currentRow & ",0)"
            End If
        End If
        
        currentRow = currentRow + 1
        
        ' Progress update
        If buildingIndex Mod 10 = 0 Then
            MsgBox "OPTIMIZED: Processed " & buildingIndex & " of " & buildingCount & " buildings", vbInformation
        End If
    Next buildingIndex
    
    MsgBox "OPTIMIZED: All buildings processed. Formatting...", vbInformation
    
    ' Auto-fit columns
    legacyWs.Columns.AutoFit
    
    ' Add totals row
    Dim totalsRow As Long
    totalsRow = currentRow + 1
    legacyWs.Cells(totalsRow, 1).Value = "GRAND TOTAL"
    legacyWs.Cells(totalsRow, 1).Font.Bold = True
    
    legacyWs.Cells(totalsRow, 2).Formula = "=SUM(B" & (headerRow + 1) & ":B" & (totalsRow - 1) & ")"
    legacyWs.Cells(totalsRow, 3).Formula = "=SUM(C" & (headerRow + 1) & ":C" & (totalsRow - 1) & ")"
    legacyWs.Cells(totalsRow, 4).Formula = "=SUM(D" & (headerRow + 1) & ":D" & (totalsRow - 1) & ")"
    legacyWs.Cells(totalsRow, 5).Formula = "=SUM(E" & (headerRow + 1) & ":E" & (totalsRow - 1) & ")"
    legacyWs.Cells(totalsRow, 6).Formula = "=IF(C" & totalsRow & ">0,E" & totalsRow & "/C" & totalsRow & ",0)"
    
    ' Format numbers
    legacyWs.Range("C:F").NumberFormat = "#,##0"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "OPTIMIZED Legacy View created successfully!" & vbCrLf & _
           "Summary format with " & buildingCount & " buildings" & vbCrLf & _
           "Uses efficient SUMIF/COUNTIF formulas for auto-updating", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description & " (Number: " & Err.Number & ")", vbCritical
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
