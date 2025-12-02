' BankScheduleSanitizer_Simple.bas
' Simplified VBA version focusing on core functionality
' Use this for easier implementation and customization

Option Explicit

' Main processing function - run this manually or on workbook open
Sub ProcessBankScheduleSimple()
    On Error GoTo ErrorHandler
    
    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    
    ' Check if source sheet exists
    If Not SheetExists("Bank Schedule") Then
        MsgBox "Please ensure the source sheet is named 'Bank Schedule'", vbCritical
        Exit Sub
    End If
    
    ' Create Units sheet
    Call CreateUnitsSheetSimple
    
    ' Create Buildings sheet
    Call CreateBuildingsSheetSimple
    
    Application.ScreenUpdating = True
    
    MsgBox "Processing complete!" & vbCrLf & _
           "• Units sheet created with filtered data" & vbCrLf & _
           "• Buildings sheet created with summary formulas" & vbCrLf & _
           "• Cap Valn values need manual setup (see guide)", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Create Units sheet with unit data only
Sub CreateUnitsSheetSimple()
    Dim sourceWS As Worksheet
    Dim unitsWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, sourceRow As Long, targetRow As Long
    
    Set sourceWS = Worksheets("Bank Schedule")
    
    ' Delete and recreate Units sheet
    If SheetExists("Units") Then
        Application.DisplayAlerts = False
        Worksheets("Units").Delete
        Application.DisplayAlerts = True
    End If
    
    Set unitsWS = Worksheets.Add
    unitsWS.Name = "Units"
    
    ' Find data boundaries
    lastRow = sourceWS.Cells(sourceWS.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWS.Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' Copy headers (row 3 from source to row 1 in Units)
    sourceWS.Range(sourceWS.Cells(3, 1), sourceWS.Cells(3, lastCol)).Copy
    unitsWS.Cells(1, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Find key columns
    Dim unitDemiseCol As Long, propertyCol As Long
    unitDemiseCol = FindColumn(unitsWS, "Unit Demise", 1)
    propertyCol = FindColumn(unitsWS, "Property", 1)
    
    If unitDemiseCol = 0 Or propertyCol = 0 Then
        MsgBox "Could not find Unit Demise or Property columns", vbCritical
        Exit Sub
    End If
    
    ' Copy rows with unit data
    targetRow = 2
    For sourceRow = 4 To lastRow
        ' Check if this row has unit data
        If Len(Trim(sourceWS.Cells(sourceRow, unitDemiseCol).Value)) > 0 And _
           Len(Trim(sourceWS.Cells(sourceRow, propertyCol).Value)) > 0 Then
            
            ' Copy entire row
            sourceWS.Range(sourceWS.Cells(sourceRow, 1), sourceWS.Cells(sourceRow, lastCol)).Copy
            unitsWS.Cells(targetRow, 1).PasteSpecial xlPasteAll
            targetRow = targetRow + 1
        End If
    Next sourceRow
    
    Application.CutCopyMode = False
    
    ' Format headers and add autofilter
    With unitsWS.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .AutoFilter
    End With
    
    ' Autofit columns
    unitsWS.Columns.AutoFit
End Sub

' Create Buildings summary sheet
Sub CreateBuildingsSheetSimple()
    Dim unitsWS As Worksheet
    Dim buildingsWS As Worksheet
    Dim buildingNames As Collection
    Dim lastRow As Long
    Dim i As Long
    
    If Not SheetExists("Units") Then
        MsgBox "Units sheet not found. Please run Units creation first.", vbCritical
        Exit Sub
    End If
    
    Set unitsWS = Worksheets("Units")
    
    ' Delete and recreate Buildings sheet
    If SheetExists("Buildings") Then
        Application.DisplayAlerts = False
        Worksheets("Buildings").Delete
        Application.DisplayAlerts = True
    End If
    
    Set buildingsWS = Worksheets.Add
    buildingsWS.Name = "Buildings"
    
    ' Create headers
    Dim headers As Variant
    headers = Array("Building", "Net Area", "Rent PA (£)", "2023 ERV (£)", "2024 ERV (£)", "Cap Valn (£)")
    
    For i = 0 To UBound(headers)
        buildingsWS.Cells(1, i + 1).Value = headers(i)
        buildingsWS.Cells(1, i + 1).Font.Bold = True
    Next i
    
    ' Get unique building names
    Set buildingNames = GetUniqueBuildings(unitsWS)
    
    ' Create formulas for each building
    Dim buildingName As Variant
    Dim row As Long
    row = 2
    
    For Each buildingName In buildingNames
        buildingsWS.Cells(row, 1).Value = buildingName
        
        ' Create SUMIF formulas (adjust column references as needed)
        buildingsWS.Cells(row, 2).Formula = CreateSumIfFormula("Net Area", buildingName)
        buildingsWS.Cells(row, 3).Formula = CreateSumIfFormula("Rent PA", buildingName)
        buildingsWS.Cells(row, 4).Formula = CreateSumIfFormula("2023 ERV", buildingName)
        buildingsWS.Cells(row, 5).Formula = CreateSumIfFormula("2024 ERV", buildingName)
        buildingsWS.Cells(row, 6).Value = "Manual setup needed"
        
        row = row + 1
    Next buildingName
    
    ' Format the sheet
    With buildingsWS.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .AutoFilter
    End With
    
    buildingsWS.Columns.AutoFit
End Sub

' Helper functions
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function FindColumn(ws As Worksheet, headerName As String, headerRow As Long) As Long
    Dim i As Long
    For i = 1 To ws.Cells(headerRow, Columns.Count).End(xlToLeft).Column
        If InStr(1, ws.Cells(headerRow, i).Value, headerName, vbTextCompare) > 0 Then
            FindColumn = i
            Exit Function
        End If
    Next i
    FindColumn = 0
End Function

Function GetUniqueBuildings(unitsWS As Worksheet) As Collection
    Dim buildings As Collection
    Dim propertyCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim building As String
    
    Set buildings = New Collection
    propertyCol = FindColumn(unitsWS, "Property", 1)
    lastRow = unitsWS.Cells(unitsWS.Rows.Count, propertyCol).End(xlUp).Row
    
    For i = 2 To lastRow
        building = Trim(unitsWS.Cells(i, propertyCol).Value)
        If Len(building) > 0 Then
            On Error Resume Next
            buildings.Add building, building
            On Error GoTo 0
        End If
    Next i
    
    Set GetUniqueBuildings = buildings
End Function

Function CreateSumIfFormula(columnName As String, buildingName As String) As String
    Dim propertyCol As String, dataCol As String
    Dim unitsWS As Worksheet
    Set unitsWS = Worksheets("Units")
    
    propertyCol = ColumnLetter(FindColumn(unitsWS, "Property", 1))
    dataCol = ColumnLetter(FindColumn(unitsWS, columnName, 1))
    
    If dataCol = "" Then
        CreateSumIfFormula = "0"
    Else
        CreateSumIfFormula = "=SUMIF(Units!" & propertyCol & ":" & propertyCol & _
                           ",""" & buildingName & """,Units!" & dataCol & ":" & dataCol & ")"
    End If
End Function

Function ColumnLetter(colNum As Long) As String
    If colNum = 0 Then
        ColumnLetter = ""
    Else
        ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
    End If
End Function
