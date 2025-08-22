Attribute VB_Name = "Module3"
Sub MacroSumSpendature()
    '
    ' Purpose: Create sum formulas for YearSpendatures sheet using Table9 and Table10 data
    ' Author: Todd Martin
    ' Date: 8/4/2025
    ' Dependencies: Requires Table9 and Table10 with monthly columns
    '

    ' Declare variables
    Dim ws As Worksheet
    Dim months As Variant
    Dim col As Integer
    Dim i As Integer

    ' Error handling
    On Error GoTo ErrorHandler

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False

    ' Set worksheet reference (assuming active sheet is YearSpendatures)
    Set ws = ActiveSheet

    ' Define month names array
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")

    ' === SECTION 1: Create monthly sum formulas in Row 5 (Table9 sums) ===
    For i = 0 To 11  ' 12 months (0-11)
        col = 4 + i  ' Start at column D (4) through O (15)
        ws.Cells(5, col).Formula = "=SUM(Table9[" & months(i) & "])"
    Next i

    ' Row 5, Column P: Sum of columns D5:O5
    ws.Range("P5").FormulaR1C1 = "=SUM(R[0]C[-12]:R[0]C[-1])"

    ' === SECTION 2: Create monthly sum formulas in Row 25 (Table10 sums) ===
    For i = 0 To 11  ' 12 months (0-11)
        col = 4 + i  ' Start at column D (4) through O (15)
        ws.Cells(25, col).Formula = "=SUM(Table10[" & months(i) & "])"
    Next i

    ' Row 25, Column P: Sum of columns D25:O25
    ws.Range("P25").FormulaR1C1 = "=SUM(R[0]C[-12]:R[0]C[-1])"

    ' === SECTION 3: Create yearly sum formulas in Column P (Rows 2-4, 7-24) ===
    ' Rows 2-4: Table9 yearly sums
    For i = 2 To 4
        ws.Cells(i, 16).Formula = "=SUM(Table9[@[January]:[December]])"
    Next i

    ' Rows 7-24: Table10 yearly sums
    For i = 7 To 24
        ws.Cells(i, 16).Formula = "=SUM(Table10[@[January]:[December]])"
    Next i

    ' === SECTION 4: Create difference formulas in Row 26 (Row 5 minus Row 25) ===
    For i = 0 To 12  ' Columns D through P (0-12 represents the 13 columns)
        col = 4 + i  ' Start at column D (4) through P (16)
        ws.Cells(26, col).FormulaR1C1 = "=R[-21]C-R[-1]C"  ' Row 5 minus Row 25
    Next i

    ' Return to cell A1
    ws.Range("A1").Select

    ' Show completion message
    MsgBox "Sum formulas have been created successfully!" & vbCrLf & _
           "• Row 5: Monthly sums from Table9" & vbCrLf & _
           "• Row 25: Monthly sums from Table10" & vbCrLf & _
           "• Column P: Yearly totals" & vbCrLf & _
           "• Row 26: Differences (Row 5 - Row 25)", _
           vbInformation, "Formulas Created"

    Call expenditures_Percentage
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "An error occurred while creating sum formulas: " & Err.description, _
           vbCritical, "Error in MacroSumSpendature"

CleanUp:
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False

    ' Clear object variables
    Set ws = Nothing

End Sub
Sub expenditures_Percentage()
    '
    ' expenditures_Percentage Macro
    '
    ' Purpose: Calculate expense percentages compared to income in YearSpendatures sheet
    ' Formula: Expenses (Row 25) divided by Income (Row 5) with error handling
    ' Location: Row 27, Columns D through P (monthly and yearly totals)
    '
    ' Created by: Todd Martin
    ' Date: August 21, 2025
    '
    ' Dependencies:
    ' - YearSpendatures worksheet must exist
    ' - Row 5 must contain income data (columns D-P)
    ' - Row 25 must contain expense data (columns D-P)
    '
    
    ' Declare variables
    Dim ws As Worksheet
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False
    
    ' Set worksheet reference
    Set ws = Sheets("YearSpendatures")
    
    ' === SECTION 1: Create percentage formula in starting cell (D27) ===
    ' Formula explanation: Row 25 (expenses) ÷ Row 5 (income)
    ' IFERROR handles division by zero, returning 0% instead of error
    ws.Range("D27").FormulaR1C1 = "=IFERROR(R[-2]C/R[-22]C,0)"
    
    ' === SECTION 2: Copy formula across all monthly columns (E27:P27) ===
    ' Copy the formula from D27 to columns E through P
    ws.Range("D27").Copy
    ws.Range("E27:P27").PasteSpecial xlPasteFormulas
    
    ' === SECTION 3: Format all percentage cells as percentages ===
    ' Apply percentage format with 2 decimal places (e.g., 75.25%)
    ws.Range("D27:P27").NumberFormat = "0.00%"
    
    ' === SECTION 4: Clean up and return to starting position ===
    ' Clear clipboard and return to cell A1
    Application.CutCopyMode = False
    ws.Range("A1").Select
    
    ' Show completion message
    MsgBox "Expenditure percentage calculations completed successfully!" & vbCrLf & _
           "• Row 27: Expense percentages (Row 25 ÷ Row 5)" & vbCrLf & _
           "• Columns D-P: Monthly and yearly percentages" & vbCrLf & _
           "• Format: Percentage with 2 decimal places" & vbCrLf & _
           "• Error handling: Shows 0% for division by zero", _
           vbInformation, "Percentage Calculations Complete"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "An error occurred while creating percentage formulas: " & Err.description, _
           vbCritical, "Error in expenditures_Percentage"
    
CleanUp:
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    
    ' Clear object variables
    Set ws = Nothing
    
End Sub

' === ENHANCED VERSION WITH STATIC VALUE CONVERSION ===
Sub MacroSumSpendature_WithStaticConversion()
    '
    ' Purpose: Create sum formulas and convert them to static values
    ' This version creates formulas and then converts them to values
    '

    ' Run the main formula creation
    Call MacroSumSpendature

    ' Convert to static values
    Call ConvertSpendatureFormulasToValues

End Sub

Sub ConvertSpendatureFormulasToValues()
    '
    ' Purpose: Convert all created formulas to static values
    '
    
    Dim ws As Worksheet
    Dim ranges As Variant
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    
    ' Define ranges that contain formulas to convert
    ranges = Array("D5:P5", "D25:P25", "P2:P4", "P7:P24", "D26:P26")
    
    ' Convert each range to values
    For i = 0 To UBound(ranges)
        ConvertRangeToValues ws.Range(ranges(i))
    Next i
    
    MsgBox "All formulas have been converted to static values!", _
           vbInformation, "Conversion Complete"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error converting formulas to values: " & Err.description, vbCritical
    
CleanUp:
    Application.ScreenUpdating = True
    Set ws = Nothing
    
End Sub

Private Sub ConvertRangeToValues(targetRange As Range)
    '
    ' Purpose: Convert a specific range from formulas to values
    ' Parameters: targetRange - The range to convert
    '
    
    On Error Resume Next
    
    With targetRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                      SkipBlanks:=False, Transpose:=False
    End With
    
    Application.CutCopyMode = False
    
End Sub

' === UTILITY SUBROUTINES ===
Sub ClearSpendatureFormulas()
    '
    ' Purpose: Clear all formula ranges (useful for testing)
    '
    
    Dim ws As Worksheet
    Dim ranges As Variant
    Dim i As Integer
    
    Set ws = ActiveSheet
    ranges = Array("D5:P5", "D25:P25", "P2:P4", "P7:P24", "D26:P26")
    
    For i = 0 To UBound(ranges)
        ws.Range(ranges(i)).ClearContents
    Next i
    
    MsgBox "All formula ranges have been cleared.", vbInformation
    
End Sub

Sub ValidateTableStructure()
    '
    ' Purpose: Validate that Table9 and Table10 exist with required columns
    '
    
    Dim tbl As ListObject
    Dim months As Variant
    Dim i As Integer
    Dim missingColumns As String
    
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    
    ' Check Table9
    On Error Resume Next
    Set tbl = ActiveSheet.ListObjects("Table9")
    If tbl Is Nothing Then
        MsgBox "Table9 not found!", vbCritical
        Exit Sub
    End If
    
    ' Check if all month columns exist in Table9
    For i = 0 To 11
        If tbl.ListColumns(months(i)) Is Nothing Then
            missingColumns = missingColumns & months(i) & " (Table9), "
        End If
    Next i
    
    ' Check Table10
    Set tbl = ActiveSheet.ListObjects("Table10")
    If tbl Is Nothing Then
        MsgBox "Table10 not found!", vbCritical
        Exit Sub
    End If
    
    ' Check if all month columns exist in Table10
    For i = 0 To 11
        If tbl.ListColumns(months(i)) Is Nothing Then
            missingColumns = missingColumns & months(i) & " (Table10), "
        End If
    Next i
    
    If missingColumns = "" Then
        MsgBox "Table structure validation passed! Both tables have all required month columns.", _
               vbInformation, "Validation Complete"
    Else
        MsgBox "Missing columns: " & Left(missingColumns, Len(missingColumns) - 2), _
               vbExclamation, "Validation Issues"
    End If
    
End Sub

