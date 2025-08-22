Attribute VB_Name = "Module2"
Sub SumEOYTotals()
    '
    ' Purpose: Calculate End of Year totals and add to next available column
    ' Author: Todd Martin
    ' Date: 8/10/2025
    ' Dependencies: Requires worksheets "YearSpendatures" and "Budget"
    '
    Call ShoweoyaggSheet
    
    ' Declare variables
    Dim wsEOY As Worksheet
    Dim wsBudget As Worksheet
    Dim currentYear As String
    Dim nextColumn As Long
    Dim headerText As String
    Dim dataRange As Range
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False
    
    ' Get current year
    currentYear = Year(Date)
    headerText = "EOY " & currentYear & " Totals"
    
    ' Set worksheet references (assuming current sheet is EOY sheet)
    Set wsEOY = ActiveSheet
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    
    ' Check if current year data already exists and find next available column
    nextColumn = FindNextAvailableColumn(wsEOY, currentYear)
    
    If nextColumn = -1 Then
        MsgBox "EOY totals for " & currentYear & " already exist. Process terminated.", _
               vbInformation, "Data Already Exists"
        GoTo CleanUp
    End If
    
    ' Add header in the next available column
    wsEOY.Cells(1, nextColumn).Value = headerText
    
    ' Set up the data range (B2:B26 equivalent in the new column)
    Set dataRange = wsEOY.Range(wsEOY.Cells(2, nextColumn), wsEOY.Cells(26, nextColumn))
    
    ' Step 1: Create formulas referencing YearSpendatures!P2:P26
    CreateInitialFormulas wsEOY, nextColumn
    
    ' Step 2: Convert all formulas to values
    ConvertFormulasToValues dataRange
    
    ' Step 4: Return to Budget sheet
    wsBudget.Select
    wsBudget.Range("A1").Select
    
    MsgBox "EOY totals for " & currentYear & " have been successfully calculated and added to column " & _
           Split(Cells(1, nextColumn).Address, "$")(1) & ".", vbInformation, "Process Complete"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "An error occurred while calculating EOY totals: " & Err.description, _
           vbCritical, "Error in SumEOYTotals"
    
CleanUp:
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    
    ' Clear object variables
    Set wsEOY = Nothing
    Set wsBudget = Nothing
    Set dataRange = Nothing
    
    
End Sub

Private Function FindNextAvailableColumn(ws As Worksheet, currentYear As String) As Long
    '
    ' Purpose: Find the next available column and check if current year already exists
    ' Parameters: ws - Worksheet to search
    '            currentYear - Current year to check for existing data
    ' Returns: Column number for next available column, or -1 if current year exists
    '
    
    Dim lastCol As Long
    Dim col As Long
    Dim cellValue As String
    
    ' Find the last used column in row 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Check existing headers for current year
    For col = 1 To lastCol
        cellValue = Trim(ws.Cells(1, col).Value)
        If InStr(1, cellValue, currentYear, vbTextCompare) > 0 And _
           InStr(1, cellValue, "EOY", vbTextCompare) > 0 Then
            FindNextAvailableColumn = -1  ' Current year already exists
            Exit Function
        End If
    Next col
    
    ' Return next available column
    FindNextAvailableColumn = lastCol + 1
    
End Function

Private Sub CreateInitialFormulas(ws As Worksheet, colNum As Long)
    '
    ' Purpose: Create formulas referencing YearSpendatures!P2:P26
    ' Parameters: ws - Worksheet reference
    '            colNum - Column number to place formulas
    '
    
    Dim i As Long
    
    ' Create formulas for rows 2-26 referencing YearSpendatures column P (16)
    ' Each formula references the corresponding row in YearSpendatures!P column
    For i = 2 To 26
        ws.Cells(i, colNum).Formula = "=YearSpendatures!P" & i
    Next i
    
End Sub

Private Sub CopyFormulasToRange(ws As Worksheet, colNum As Long, targetRange As Range)
    '
    ' Purpose: Handle any additional formula copying if needed
    ' Parameters: ws - Worksheet reference
    '            colNum - Column number containing formulas
    '            targetRange - Range to fill with formulas
    '
    ' Note: This function is now simplified since CreateInitialFormulas handles all rows
    
    ' The formulas are already created in CreateInitialFormulas
    ' This function is kept for compatibility but may not be needed
    ' unless there are special cases for specific rows
    
    Application.CutCopyMode = False
    
End Sub

Private Sub ConvertFormulasToValues(targetRange As Range)
    '
    ' Purpose: Convert formulas in a range to their calculated values
    ' Parameters: targetRange - The range containing formulas to convert
    '
    
    With targetRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                      SkipBlanks:=False, Transpose:=False
    End With
    
    Application.CutCopyMode = False
    
End Sub

