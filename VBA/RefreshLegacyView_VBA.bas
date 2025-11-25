Sub RefreshLegacyView()
    ' Dynamic Legacy View Rebuilder
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
    Dim buildingDict As Object
    Dim buildingName As String
    Dim key As Variant
    Dim unitRows As Collection
    Dim i As Long
    Dim startRow As Long, endRow As Long
    
    ' Get worksheets
    Set unitsWs = ThisWorkbook.Worksheets("Units")
    
    ' Clear existing Legacy View content (keep first 3 instruction rows)
    Set legacyWs = ThisWorkbook.Worksheets("Legacy View")
    lastRow = legacyWs.Cells(legacyWs.Rows.Count, 1).End(xlUp).Row
    If lastRow > 3 Then
        legacyWs.Range("A4:ZZ" & lastRow).Clear
    End If
    
    ' Set up headers
    headerRow = 4
    For col = 2 To unitsWs.Cells(1, unitsWs.Columns.Count).End(xlToLeft).Column
        If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
            legacyWs.Cells(headerRow, col - 1).Value = unitsWs.Cells(1, col).Value
            legacyWs.Cells(headerRow, col - 1).Font.Bold = True
        End If
    Next col
    
    ' Group units by building
    Set buildingDict = CreateObject("Scripting.Dictionary")
    lastRow = unitsWs.Cells(unitsWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        buildingName = Trim(CStr(unitsWs.Cells(i, 1).Value))
        If buildingName <> "" Then
            If Not buildingDict.Exists(buildingName) Then
                Set buildingDict(buildingName) = New Collection
            End If
            buildingDict(buildingName).Add i
        End If
    Next i
    
    ' Build hierarchical structure
    currentRow = headerRow + 2 ' Empty row after headers
    
    For Each key In buildingDict.Keys
        Set unitRows = buildingDict(key)
        
        ' Building header
        legacyWs.Cells(currentRow, 1).Value = key
        legacyWs.Cells(currentRow, 1).Font.Bold = True
        legacyWs.Cells(currentRow, 1).Font.Size = 12
        currentRow = currentRow + 2 ' Empty row after building header
        
        ' Unit rows
        startRow = currentRow
        For i = 1 To unitRows.Count
            For col = 2 To unitsWs.Cells(1, unitsWs.Columns.Count).End(xlToLeft).Column
                If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
                    legacyWs.Cells(currentRow, col - 1).Formula = "=Units!" & _
                        ConvertToColumnLetter(col) & unitRows(i)
                End If
            Next col
            currentRow = currentRow + 1
        Next i
        endRow = currentRow - 1
        
        ' Empty row + summary row
        currentRow = currentRow + 1
        legacyWs.Cells(currentRow, 1).Value = key & " - TOTAL"
        legacyWs.Cells(currentRow, 1).Font.Bold = True
        legacyWs.Cells(currentRow, 1).Font.Italic = True
        
        ' Add SUM formulas for numeric columns
        For col = 2 To unitsWs.Cells(1, unitsWs.Columns.Count).End(xlToLeft).Column
            If Not IsEmpty(unitsWs.Cells(1, col).Value) Then
                If IsNumericColumn(unitsWs.Cells(1, col).Value) Then
                    If startRow <= endRow Then
                        legacyWs.Cells(currentRow, col - 1).Formula = "=SUM(" & _
                            ConvertToColumnLetter(col - 1) & startRow & ":" & _
                            ConvertToColumnLetter(col - 1) & endRow & ")"
                        legacyWs.Cells(currentRow, col - 1).Font.Bold = True
                    End If
                End If
            End If
        Next col
        
        currentRow = currentRow + 2 ' Empty row after summary
    Next key
    
    ' Auto-fit columns
    legacyWs.Columns.AutoFit
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Legacy View refreshed successfully! Found " & buildingDict.Count & " buildings.", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error refreshing Legacy View: " & Err.Description, vbCritical
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

Sub Auto_Open()
    ' Set up keyboard shortcut when workbook opens
    Application.OnKey "^+R", "RefreshLegacyView"
End Sub