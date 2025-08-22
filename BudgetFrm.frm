VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BudgetFrm 
   Caption         =   "Budget Form"
   ClientHeight    =   9012.001
   ClientLeft      =   384
   ClientTop       =   1968
   ClientWidth     =   12360
   OleObjectBlob   =   "BudgetFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BudgetFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

' Set the form size first
Me.Height = 475.8
Me.Width = 645

' Optional: Center the form on screen
Me.StartUpPosition = 0 ' Manual positioning
Me.Left = (Application.Width - Me.Width) / 2
Me.Top = (Application.Height - Me.Height) / 2

With Me.month_cbo ' Refers to your ComboBox named 'month_cbo'
    .Clear ' Clears any existing items in the ComboBox

    ' Add each month name to the ComboBox
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With

' Call the savings calculation when the form initializes
Call CalculateSavingsGoalStatus
Copy

End Sub

Private Sub month_cbo_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Me.month_cbo.ListIndex >= 0 Then
    ThisWorkbook.Sheets("Budget").Range("A1").Value = Me.month_cbo.Value
    Debug.Print "Month updated via Exit event: " & Me.month_cbo.Value
End If
End Sub

Private Sub exitform_btn_Click() ' This code will execute when the 'exitform_btn' button is clicked.
Unload Me ' Unloads (closes) the current UserForm.

End Sub

Private Sub submit_btn_Click() ' This subroutine orchestrates the saving, updating, and transferring of data.

On Error GoTo ErrorHandler

    ' Save data to the Budget sheet
    Call saveDataToSheet_Click
    
    ' Sync donation details to YearSpendatures sheet
    Call SyncDonationDetailsToYearSpendatures
    
    ' Transfer budget data to the YearSpendatures sheet
    Call TransferBudgetToYearSpendatures
    
    ' >>> NEW: Call the aggregate updater after donations are saved <<<
    Call UpdateDonationAggregate
    
    ' >>> NEW: Convert all formulas to static values <<<
    ' Call ConvertFormulasToStaticValues
    
    ' call refresh pivot table
     Call RefreshExpensePivotTableSilent
     
    ' Display completion message
    MsgBox "All data operations completed successfully and converted to static values!", vbInformation
    
    ' Close the Budget Form
    Unload Me
    Exit Sub
Copy

ErrorHandler: MsgBox "An error occurred during submission: " & Err.description, vbCritical
End Sub
' Module: SumYearSpendatures

Sub SumYearSpendatures()
    '
    ' Purpose: Calculate sums for YearSpendatures sheet and update Budget sheet
    ' Author: [Your Name]
    ' Date: [Current Date]
    '
    
    ' Declare variables
    Dim wsSource As Worksheet
    Dim wsBudget As Worksheet
    Dim col As Integer
    Dim sumTop As Double
    Dim sumBottom As Double
    Dim totalSum As Double
    Dim diff As Double
    Dim monthName As String
    Dim i As Long
    Dim matchRow As Long
    Dim foundMatch As Boolean
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet references
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    
    ' === Part 1: Call to sum the Yearly Spendatures ===
    Call MacroSumSpendature
    
    ' === Part 2: Copy P7 to Budget!N17 ===
    wsBudget.Range("N17").Value = wsSource.Range("P7").Value
    
    ' === Part 3: Budget month data processing ===
    monthName = Trim(wsBudget.Range("A1").Value)
    foundMatch = False
    
    ' Search L3:L14 for monthName
    For i = 3 To 14
        If Trim(wsBudget.Cells(i, "L").Value) = Trim(monthName) Then
            matchRow = i + 3 ' Map L3 ? row 6, L4 ? 7, etc.
            wsBudget.Cells(matchRow, "M").Value = wsBudget.Range("E6").Value
            wsBudget.Cells(matchRow, "N").Value = wsBudget.Range("E31").Value
            foundMatch = True
            Exit For
        End If
    Next i
    
    ' Calculate O6:O14 = M - N
    For i = 6 To 14
        wsBudget.Cells(i, "O").Value = wsBudget.Cells(i, "M").Value - wsBudget.Cells(i, "N").Value
    Next i
    
    ' Calculate sums for Budget sheet
    wsBudget.Range("M15").Value = Application.WorksheetFunction.Sum(wsBudget.Range("M6:M14"))
    wsBudget.Range("N15").Value = Application.WorksheetFunction.Sum(wsBudget.Range("N6:N14"))
    wsBudget.Range("M16").Value = wsBudget.Range("M15").Value + wsBudget.Range("N15").Value
    wsBudget.Range("O15").Value = Application.WorksheetFunction.Sum(wsBudget.Range("O6:O14"))
    
    ' Fixed: Correct reference to wsSource instead of undefined wsYearSpendatures
    wsBudget.Range("N17").Value = wsSource.Range("P7").Value
    
    ' === Part 4: Convert calculated values to static values (REMOVED FOR NOW) ===
    ' Comment out the conversion temporarily to see if calculations are working
    ' ConvertCalculationsToValues wsSource, wsBudget
    
    ' Show completion message with calculation summary
    MsgBox "YearSpendatures calculations completed!" & vbCrLf & _
           "Check rows 5, 25, and 26 for calculated values." & vbCrLf & _
           "Debug info printed to Immediate Window (Ctrl+G to view).", _
           vbInformation, "Process Complete"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description, vbCritical, "Error in SumYearSpendatures"
    
CleanUp:
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Clear object variables
    Set wsSource = Nothing
    Set wsBudget = Nothing
    
End Sub

Private Sub ConvertCalculationsToValues(wsSource As Worksheet, wsBudget As Worksheet)
    '
    ' Purpose: Convert all calculated values to static values
    ' Parameters: wsSource - YearSpendatures worksheet
    '            wsBudget - Budget worksheet
    '
    
    On Error Resume Next
    
    ' Convert YearSpendatures calculated ranges to values
    ' Rows 5, 25, 26 for columns D through P
    Dim calcRange As Range
    
    ' Row 5 (D5:P5)
    Set calcRange = wsSource.Range("D5:P5")
    calcRange.Copy
    calcRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    
    ' Row 25 (D25:P25)
    Set calcRange = wsSource.Range("D25:P25")
    calcRange.Copy
    calcRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    
    ' Row 26 (D26:P26)
    Set calcRange = wsSource.Range("D26:P26")
    calcRange.Copy
    calcRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    
    ' Convert Budget sheet calculated values to static
    ' Columns M, N, O for rows 6-17
    Set calcRange = wsBudget.Range("M6:O17")
    calcRange.Copy
    calcRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
    
End Sub

' === SEPARATE SUBROUTINE TO CONVERT TO STATIC VALUES ===
Sub ConvertYearSpendaturesToStaticValues()
    '
    ' Purpose: Convert all calculated values to static values
    ' Run this AFTER SumYearSpendatures() if you want static values
    '
    
    Dim wsSource As Worksheet
    Dim wsBudget As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    
    Application.ScreenUpdating = False
    
    ConvertCalculationsToValues wsSource, wsBudget
    
    MsgBox "All calculated values have been converted to static values!", _
           vbInformation, "Conversion Complete"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error converting to static values: " & Err.description, vbCritical
    
CleanUp:
    Application.ScreenUpdating = True
    Set wsSource = Nothing
    Set wsBudget = Nothing
    
End Sub

' === DIAGNOSTIC SUBROUTINE ===
Sub DiagnoseYearSpendaturesData()
    '
    ' Purpose: Diagnose what data exists in the calculation ranges
    '
    
    Dim wsSource As Worksheet
    Dim col As Integer
    Dim r As Integer
    Dim cellValue As Variant
    Dim debugMsg As String
    
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    
    debugMsg = "Data Analysis for YearSpendatures:" & vbCrLf & vbCrLf
    
    ' Check column D (first column) for sample data
    col = 4 ' Column D
    debugMsg = debugMsg & "Column D Analysis:" & vbCrLf
    debugMsg = debugMsg & "Rows 2-4 (Budget):" & vbCrLf
    
    For r = 2 To 4
        cellValue = wsSource.Cells(r, col).Value
        debugMsg = debugMsg & "  Row " & r & ": " & cellValue & " (" & TypeName(cellValue) & ")" & vbCrLf
    Next r
    
    debugMsg = debugMsg & "Rows 7-10 (Sample Actual):" & vbCrLf
    For r = 7 To 10
        cellValue = wsSource.Cells(r, col).Value
        debugMsg = debugMsg & "  Row " & r & ": " & cellValue & " (" & TypeName(cellValue) & ")" & vbCrLf
    Next r
    
    MsgBox debugMsg, vbInformation, "Data Diagnosis"
    
End Sub

' === ALTERNATIVE VERSION WITH DIFFERENT ROW 25 CALCULATION ===
' Use this version if row 25 should have a different calculation

Sub SumYearSpendatures_Alternative()
    '
    ' Alternative version with different row 25 calculation
    ' Use this if row 25 should be calculated differently
    '
    
    Dim wsSource As Worksheet
    Dim wsBudget As Worksheet
    Dim col As Integer
    Dim sumTop As Double
    Dim sumBottom As Double
    Dim totalSum As Double
    Dim actualSpent As Double
    Dim diff As Double
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    
    For col = 4 To 16 ' Columns D to P
        
        ' Row 5: Total budget (rows 2-4 + rows 7-24)
        sumTop = Application.WorksheetFunction.Sum(wsSource.Range(wsSource.Cells(2, col), wsSource.Cells(4, col)))
        sumBottom = Application.WorksheetFunction.Sum(wsSource.Range(wsSource.Cells(7, col), wsSource.Cells(24, col)))
        totalSum = sumTop + sumBottom
        wsSource.Cells(5, col).Value = totalSum
        
        ' Row 25: Actual spent amount (only bottom section)
        actualSpent = sumBottom
        wsSource.Cells(25, col).Value = actualSpent
        
        ' Row 26: Difference (Budget - Actual)
        diff = wsSource.Cells(5, col).Value - wsSource.Cells(25, col).Value
        wsSource.Cells(26, col).Value = diff
        
    Next col
    
    ' Rest of the Budget sheet processing...
    ' [Include the same Budget sheet code as above]
    
    ' Convert to static values
    ConvertCalculationsToValues wsSource, wsBudget
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.description, vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub ' Corrected VBA for YearSpendatures sheet to fix row 5, 25, and 26 calculations

Sub ProcessYearSpendaturesSums()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("YearSpendatures")

    Dim col As Long
    Dim sumTop As Double
    Dim sumBottom As Double
    Dim difference As Double

    For col = 4 To 15 ' Columns D (4) to O (15)

        ' Sum D2:D4 ... O2:O4
        sumTop = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, col), ws.Cells(4, col)))

        ' Sum D7:D24 ... O7:O24
        sumBottom = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(7, col), ws.Cells(24, col)))

        ' Write sum to row 5 and 25
        ws.Cells(5, col).Value = sumTop
        ws.Cells(25, col).Value = sumBottom

        ' Row 26 = Row 5 - Row 25
        difference = sumTop - sumBottom
        ws.Cells(26, col).Value = difference

    Next col

End Sub

' --- Subroutine from BudgetFrm --- ' This subroutine transfers data from the UserForm to the Budget worksheet.
Private Sub saveDataToSheet_Click()
    
    Dim ws As Worksheet
    Dim cellL As Range
    Dim targetValue As Variant
    Dim sourceE6 As Variant
    Dim sourceE31 As Variant


On Error GoTo ErrorHandler

Set ws = ThisWorkbook.Sheets("Budget")

' Transfer the Month combo box - ENSURE this is set first
    If Me.month_cbo.ListIndex >= 0 Then
        ws.Range("A1").Value = Me.month_cbo.Value
        ' Debug print to verify month is set
        Debug.Print "Month set in Budget A1: " & ws.Range("A1").Value
    Else
        MsgBox "Please select a month from the dropdown list.", vbExclamation, "Month Selection Required"
        Exit Sub
    End If

' Transfer Income fields - Use Val() for numeric fields
    ws.Range("D3").Value = Val(Me.income1_txt.Value)
    ws.Range("D4").Value = Val(Me.income2_txt.Value)
    ws.Range("D5").Value = Val(Me.otherincome_txt.Value)

' Transfer Savings/Investments/Donations fields - Use Val() for numeric fields
    ws.Range("D8").Value = Val(Me.savings_txt.Value)
    ws.Range("D9").Value = Val(Me.donations_txt.Value)
    ws.Range("D10").Value = Val(Me.Investments_txt.Value)

' Transfer Housing Expenses fields - Use Val() for numeric fields
    ws.Range("D12").Value = Val(Me.mrtg_txt.Value)
    ws.Range("D13").Value = Val(Me.utility_txt.Value)
    ws.Range("D14").Value = Val(Me.otherexp_txt.Value)

' Transfer Transportation fields - Use Val() for numeric fields
    ws.Range("D16").Value = Val(Me.trans_txt.Value)
    ws.Range("D17").Value = Val(Me.carIns_txt.Value)
    ws.Range("D18").Value = Val(Me.other_trans_txt.Value)

' Transfer Food fields - Use Val() for numeric fields
    ws.Range("D20").Value = Val(Me.groceries_txt.Value)
    ws.Range("D21").Value = Val(Me.takeout_txt.Value)
    ws.Range("D22").Value = Val(Me.otherfood_txt.Value)

' Transfer Credit/Medical fields - Use Val() for numeric fields
    ws.Range("D24").Value = Val(Me.credit_txt.Value)
    ws.Range("D25").Value = Val(Me.medical_txt.Value)
    ws.Range("D26").Value = Val(Me.othercredit_txt.Value)

' Transfer Other Expenses fields - Use Val() for numeric fields
    ws.Range("D28").Value = Val(Me.cothing_txt.Value)
    ws.Range("D29").Value = Val(Me.entertainment_txt.Value)
    ws.Range("D30").Value = Val(Me.other_txt.Value)

' --- Correction: Get the target month from cell A1, not R1 ---
' This is the month value that was just set from the UserForm.
    targetValue = ws.Range("A1").Value
    sourceE6 = ws.Range("E6").Value
    sourceE31 = ws.Range("E31").Value

    For Each cellL In ws.Range("L3:L14")
        ' --- Correction: Use LCase for case-insensitive comparison ---
        If LCase(Trim(cellL.Value)) = LCase(Trim(targetValue)) Then
            ' Store calculated values directly instead of formulas
            ws.Range("M" & cellL.Row).Value = Val(sourceE6)
            ws.Range("N" & cellL.Row).Value = Val(sourceE31)
            ws.Range("O" & cellL.Row).Value = Val(ws.Range("M" & cellL.Row).Value) - Val(ws.Range("N" & cellL.Row).Value)
            Exit For
        End If
    Next cellL
        
Exit Sub
Copy
Call Cell_Center
Call YS_SumTotals
Call ValueStatic
Call BudgetTotals
Call SavingSum
ErrorHandler:
    MsgBox "An error occurred in saveDataToSheet_Click: " & Err.description, vbCritical
End Sub

Private Sub Workbook_Open() ' Ensure your UserForm is named BudgetFrm in the VBA Project Explorer

BudgetFrm.Show
End Sub

Private Sub TransferBudgetToYearSpendatures() ' This subroutine copies data from the 'Budget
' worksheet to the 'YearSpendatures' worksheet
' based on the month in 'Budget'!A1 and the month headers in 'YearSpendatures'!D1:P1.
' It also calculates row totals in column P for specified ranges.
' MODIFIED: Now stores calculated values directly instead of formulas ' FIXED: Removed duplicate month population in donation section

    Dim wsBudget As Worksheet
    Dim wsYearSpendatures As Worksheet
    Dim budgetMonth As String
    Dim monthColumn As Range
    Dim foundMonthColumn As Long
    Dim i As Long
    Dim rowNum As Long
    Dim calculatedValue As Double

On Error GoTo ErrorHandler

' Set the worksheets
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    Set wsYearSpendatures = ThisWorkbook.Sheets("YearSpendatures")

' Get the month from Budget!A1
budgetMonth = Trim(wsBudget.Range("A1").Value) ' Use Trim to remove any leading/trailing spaces

' If Budget!A1 is empty, default to current month
    If budgetMonth = "" Then
        budgetMonth = Format(Date, "mmmm") ' Get full month name (e.g., "July")
        ' Optionally, you could also update A1 with this default value if desired:
        wsBudget.Range("A1").Value = budgetMonth
    End If

' Debug information
Debug.Print "TransferBudgetToYearSpendatures - Processing month: " & budgetMonth

' Find the corresponding month column in YearSpendatures!D1:P1
Set monthColumn = wsYearSpendatures.Range("D1:P1").Find(What:=budgetMonth, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not monthColumn Is Nothing Then
        ' Month found, get its column number
        foundMonthColumn = monthColumn.Column
        Debug.Print "Found month column: " & foundMonthColumn & " (Column " & Split(Cells(1, foundMonthColumn).Address, "$")(1) & ")"
    
        ' Define the source cells in Budget and their corresponding destination rows in YearSpendatures
        Dim sourceCells(20) As String ' 21 cells to copy (0 to 20)
        Dim destRows(20) As Long
    
        ' Populate the arrays with source cell addresses and destination row numbers
        sourceCells(0) = "D3": destRows(0) = 2    ' Income 1
        sourceCells(1) = "D4": destRows(1) = 3    ' Income 2
        sourceCells(2) = "D5": destRows(2) = 4    ' Other Income
        sourceCells(3) = "D8": destRows(3) = 7    ' Savings
        sourceCells(4) = "D9": destRows(4) = 8    ' Donations
        sourceCells(5) = "D10": destRows(5) = 9   ' Investments
        sourceCells(6) = "D12": destRows(6) = 10  ' Mortgage/Rent
        sourceCells(7) = "D13": destRows(7) = 11  ' Utilities
        sourceCells(8) = "D14": destRows(8) = 12  ' Other Housing
        sourceCells(9) = "D16": destRows(9) = 13  ' Transportation
        sourceCells(10) = "D17": destRows(10) = 14 ' Car Insurance
        sourceCells(11) = "D18": destRows(11) = 15 ' Other Transportation
        sourceCells(12) = "D20": destRows(12) = 16 ' Groceries
        sourceCells(13) = "D21": destRows(13) = 17 ' Takeout
        sourceCells(14) = "D22": destRows(14) = 18 ' Other Food
        sourceCells(15) = "D24": destRows(15) = 19 ' Credit
        sourceCells(16) = "D25": destRows(16) = 20 ' Medical
        sourceCells(17) = "D26": destRows(17) = 21 ' Other Credit/Medical
        sourceCells(18) = "D28": destRows(18) = 22 ' Clothing
        sourceCells(19) = "D29": destRows(19) = 23 ' Entertainment
        sourceCells(20) = "D30": destRows(20) = 24 ' Other Expenses
    
        ' Loop through the defined mappings and transfer values
        For i = LBound(sourceCells) To UBound(sourceCells)
            ' Use Val() to ensure numeric values are transferred, preventing type mismatch
            wsYearSpendatures.Cells(destRows(i), foundMonthColumn).Value = Val(wsBudget.Range(sourceCells(i)).Value)
            Debug.Print "Transferred " & sourceCells(i) & " (" & wsBudget.Range(sourceCells(i)).Value & ") to row " & destRows(i) & ", column " & foundMonthColumn
        Next i
    
        ' *** MODIFIED: Calculate and store static values instead of formulas ***
        Application.Calculation = xlCalculationManual ' Temporarily disable auto-calculation for performance
        
        ' Calculate totals for Income rows (P2:P4) and store as static values
        For rowNum = 2 To 4
            calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("D" & rowNum & ":O" & rowNum))
            wsYearSpendatures.Cells(rowNum, "P").Value = calculatedValue
            Debug.Print "Calculated static value for P" & rowNum & ": " & calculatedValue
        Next rowNum
        
        ' Calculate total for Income summary (P5 - sum of P2:P4) as static value
        calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("P2:P4"))
        wsYearSpendatures.Cells(5, "P").Value = calculatedValue
        Debug.Print "Calculated static value for P5: " & calculatedValue
        
        ' Calculate totals for Expenses rows (P7:P24) and store as static values
        For rowNum = 7 To 24
            calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("D" & rowNum & ":O" & rowNum))
            wsYearSpendatures.Cells(rowNum, "P").Value = calculatedValue
            Debug.Print "Calculated static value for P" & rowNum & ": " & calculatedValue
        Next rowNum
        
        ' Calculate totals for summary rows as static values
        calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("P7:P24"))
        wsYearSpendatures.Cells(25, "P").Value = calculatedValue
        Debug.Print "Calculated static value for P25: " & calculatedValue
        
        calculatedValue = wsYearSpendatures.Cells(5, "P").Value - wsYearSpendatures.Cells(25, "P").Value
        wsYearSpendatures.Cells(26, "P").Value = calculatedValue
        Debug.Print "Calculated static value for P26: " & calculatedValue
        
        Application.Calculation = xlCalculationAutomatic ' Re-enable auto-calculation
        
        Debug.Print "Budget data and static calculations completed for " & budgetMonth
        
    Else
        MsgBox "Month '" & budgetMonth & "' not found in YearSpendatures headers (D1:P1). Data not transferred.", vbExclamation
        Debug.Print "ERROR: Month '" & budgetMonth & "' not found in headers D1:P1"
    End If

    Exit Sub ' Exit the sub if no error
Copy

ErrorHandler:     Application.Calculation = xlCalculationAutomatic ' Ensure auto-calculation is restored
    MsgBox "An error occurred in TransferBudgetToYearSpendatures: " & Err.description, vbCritical
    Debug.Print "ERROR in TransferBudgetToYearSpendatures: " & Err.description
End Sub
' Run this once to set up all the calculated values as static data instead of formulas
Sub InitializeYearSpendaturesStaticValues()
    Dim wsYearSpendatures As Worksheet
    Dim rowNum As Long
    Dim calculatedValue As Double

On Error GoTo ErrorHandler

Set wsYearSpendatures = ThisWorkbook.Sheets("YearSpendatures")

Application.Calculation = xlCalculationManual

' Calculate and store static values for Income rows (P2:P4)
    For rowNum = 2 To 4
        calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("D" & rowNum & ":O" & rowNum))
        wsYearSpendatures.Cells(rowNum, "P").Value = calculatedValue
    Next rowNum
    
    ' Calculate and store static value for total Income (P5)
    calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("P2:P4"))
    wsYearSpendatures.Cells(5, "P").Value = calculatedValue
    
    ' Calculate and store static values for Expense rows (P7:P24)
    For rowNum = 7 To 24
        calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("D" & rowNum & ":O" & rowNum))
        wsYearSpendatures.Cells(rowNum, "P").Value = calculatedValue
    Next rowNum

' Calculate and store static values for summary
    calculatedValue = Application.WorksheetFunction.Sum(wsYearSpendatures.Range("P7:P24"))
    wsYearSpendatures.Cells(25, "P").Value = calculatedValue  ' Total Expenses
    
    calculatedValue = wsYearSpendatures.Cells(5, "P").Value - wsYearSpendatures.Cells(25, "P").Value
    wsYearSpendatures.Cells(26, "P").Value = calculatedValue  ' Net (Income - Expenses)
    
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "YearSpendatures static values initialized successfully!", vbInformation
    Exit Sub
    Copy

ErrorHandler: Application.Calculation = xlCalculationAutomatic
MsgBox "Error initializing static values: " & Err.description, vbCritical

End Sub

Private Sub CalculateSavingsGoalStatus() ' This subroutine calculates the savings goal status and updates UserForm labels.
' This subroutine MUST be placed in the code module of your UserForm (e.g., BudgetFrm).

    Dim savingsGoal As Double
    Dim currentSavings As Double
    Dim monthsRemaining As Double
    Dim additionalSavingsNeededPerMonth As Double
    Dim projectedEOYSavings As Double
    Dim wsBudget As Worksheet
    Dim targetMonthlySavings As Double
    Dim monthsPassed As Integer
    Dim expectedSavingsByNow As Double
    Dim remainingAmountToSave As Double

On Error GoTo ErrorHandler

' Set the worksheet
    Set wsBudget = ThisWorkbook.Sheets("Budget")

' Get the savings goal from Budget!M17 - Use Val() to prevent Type Mismatch
    savingsGoal = Val(wsBudget.Range("M17").Value)

' Get the current savings balance from Budget!N17 - Use Val() to prevent Type Mismatch
    currentSavings = Val(wsBudget.Range("N17").Value)

' Get the number of months remaining from Budget!P1 - Use Val() to prevent Type Mismatch
    monthsRemaining = Val(wsBudget.Range("P1").Value)

' Debug information
    Debug.Print "Savings Goal: " & savingsGoal
    Debug.Print "Current Savings: " & currentSavings
    Debug.Print "Months Remaining: " & monthsRemaining

' Validate data - ensure we have valid values
    If savingsGoal <= 0 Then
        Me.gsaving_lbl.caption = "$0.00"
        Me.ytd_savings.caption = "$0.00"
        Me.onTrk_lbl.caption = "No Goal Set"
        Me.onTrk_lbl.ForeColor = RGB(128, 128, 128) ' Gray color
        Me.response_lbl.caption = "Please set a savings goal in Budget sheet cell M17."
        Me.response_lbl.ForeColor = RGB(128, 128, 128)
        Me.imgHappyFace.Visible = False
        Me.imgSadFace.Visible = False
        Exit Sub
    End If

' Populate the 'gsaving_lbl' label on the UserForm with the goal
    Me.gsaving_lbl.caption = Format(savingsGoal, "$#,##0.00")

' Display current savings balance on 'ytd_savings' label
    Me.ytd_savings.caption = Format(currentSavings, "$#,##0.00")

' Ensure monthsRemaining is valid for calculations
    If monthsRemaining <= 0 Then
        monthsRemaining = 1 ' Use 1 to avoid division by zero
    End If

' Calculate the required monthly savings from the start of the year
    targetMonthlySavings = savingsGoal / 12

' Calculate the expected savings by the current point in the year
    monthsPassed = 12 - monthsRemaining
        If monthsPassed < 0 Then monthsPassed = 0 ' Ensure non-negative
        
        If monthsPassed > 0 Then
            expectedSavingsByNow = targetMonthlySavings * monthsPassed
        Else
            expectedSavingsByNow = 0 ' No months passed yet (e.g., January 1st)
        End If

    Debug.Print "Target Monthly Savings: " & targetMonthlySavings
    Debug.Print "Months Passed: " & monthsPassed
    Debug.Print "Expected Savings by Now: " & expectedSavingsByNow

' Determine if on track and display appropriate message and image
    If currentSavings >= expectedSavingsByNow Then
        ' ON TRACK - Show happy face and positive message
        Me.onTrk_lbl.caption = "On Track"
        Me.onTrk_lbl.ForeColor = RGB(0, 128, 0) ' Green color
    
        ' Show happy face image, hide sad face image
        Me.imgHappyFace.Visible = True
        Me.imgSadFace.Visible = False
    
        ' Calculate projected EOY savings based on current rate
        If monthsPassed > 0 Then
            Dim currentMonthlySavingsRate As Double
            currentMonthlySavingsRate = currentSavings / monthsPassed
            projectedEOYSavings = currentMonthlySavingsRate * 12
        Else
            ' If no months have passed, assume current savings will continue monthly
            projectedEOYSavings = currentSavings * 12
        End If

        Me.response_lbl.caption = "Congratulations, you are on track to reach your goal. At this rate you will have saved " & _
                                  Format(projectedEOYSavings, "$#,##0.00") & " by the end of the year."
        Me.response_lbl.ForeColor = RGB(0, 128, 0) ' Green for positive message
    
        Debug.Print "ON TRACK - Projected EOY Savings: " & projectedEOYSavings
    Else
        ' NOT ON TRACK - Show sad face and recommendation
        Me.onTrk_lbl.caption = "Not On Track"
        Me.onTrk_lbl.ForeColor = RGB(255, 0, 0) ' Red color
    
        ' Show sad face image, hide happy face image
        Me.imgHappyFace.Visible = False
        Me.imgSadFace.Visible = True

    ' Calculate additional savings needed per month
    remainingAmountToSave = savingsGoal - currentSavings

    If monthsRemaining > 0 Then
        additionalSavingsNeededPerMonth = remainingAmountToSave / monthsRemaining
    Else
        ' If no months remaining, need to save the full amount now
        additionalSavingsNeededPerMonth = remainingAmountToSave
    End If

    ' Ensure the additional amount is not negative
    If additionalSavingsNeededPerMonth < 0 Then
        additionalSavingsNeededPerMonth = 0
    End If

    Me.response_lbl.caption = "It is recommended to increase your savings by an additional " & _
                              Format(additionalSavingsNeededPerMonth, "$#,##0.00") & " per month."
    Me.response_lbl.ForeColor = RGB(255, 0, 0) ' Red for warning message

    Debug.Print "NOT ON TRACK - Additional Monthly Savings Needed: " & additionalSavingsNeededPerMonth
End If

Exit Sub ' Exit the sub if no error
Copy

ErrorHandler: MsgBox "An error occurred in CalculateSavingsGoalStatus: " & Err.description, vbCritical


' Hide both images in case of error
On Error Resume Next
Me.imgHappyFace.Visible = False
Me.imgSadFace.Visible = False
On Error GoTo 0
Copy

End Sub

Private Sub donations_txt_Enter()
If Me.Visible = False Then Exit Sub ' Don't run if form is still initializing


Dim frmDonation As Donation_Frm
Set frmDonation = New Donation_Frm

Set frmDonation.ParentBudgetForm = Me

frmDonation.Show
Copy

End Sub


' --- BudgetFrm Code: SyncDonationDetailsToYearSpendatures ---
Sub SyncDonationDetailsToYearSpendatures()
    Dim wsBudget As Worksheet
    Dim wsYearSpendatures As Worksheet
    Dim budgetMonth As String
    Dim donationAmount As Double
    Dim description As String
    Dim i As Long
    Dim found As Boolean
    Dim targetRow As Long
    
    On Error GoTo ErrorHandler

    Set wsBudget = ThisWorkbook.Sheets("Budget")
    Set wsYearSpendatures = ThisWorkbook.Sheets("YearSpendatures")

    budgetMonth = Trim(wsBudget.Range("A1").Value)
    donationAmount = Val(wsBudget.Range("D9").Value)
    description = Trim(wsBudget.Range("E9").Value)

    If description = "" Or IsNumeric(description) Then
        description = Trim(wsBudget.Range("F9").Value)
    End If

    If budgetMonth = "" Then
        MsgBox "No month selected. Cannot sync donation details.", vbExclamation
        Exit Sub
    End If

    Debug.Print "SyncDonationDetailsToYearSpendatures - Month: " & budgetMonth & ", Amount: " & donationAmount & ", Description: " & description


    For i = 30 To 200
        If Trim(LCase(wsYearSpendatures.Cells(i, "B").Value)) = Trim(LCase(budgetMonth)) Then
        ' Replace the comment instead of cell value
        With wsYearSpendatures.Cells(i, "E")
            If Not .Comment Is Nothing Then .Comment.Delete
            If description <> "" Then .AddComment description
        End With

        wsYearSpendatures.Cells(i, "C").Value = Date
        wsYearSpendatures.Cells(i, "D").Value = donationAmount

            found = True
            Exit For
        End If
    Next i


    If Not found Then
        MsgBox "Could not sync donation details. Month '" & budgetMonth & "' not found in YearSpendatures B30:B200.", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in SyncDonationDetailsToYearSpendatures: " & Err.description, vbCritical
End Sub
Sub SavingSum()
'
' SavingSum Macro
    
    ActiveCell.FormulaR1C1 = "=YearSpendatures!R[-10]C[2]"
    Range("N18").Select
End Sub
Sub BudgetTotals()
'
' BudgetTotals Macro
''
    Range("M15:O15").Select
    Selection.FormulaR1C1 = "=SUBTOTAL(109,R[-12]C:R[-1]C)"
    Range("A1").Select
End Sub

Sub ValueStatic()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("YearSpendatures")

    ws.Range("D5:P5").Value = ws.Range("D5:P5").Value
    ws.Range("D25:P26").Value = ws.Range("D25:P26").Value
    ws.Range("P2:P5").Value = ws.Range("P2:P5").Value
    ws.Range("P7:P24").Value = ws.Range("P7:P24").Value

    ' Avoid select
    With ThisWorkbook.Sheets("Budget")
        .Range("A1").Value = .Range("A1").Value
    End With
End Sub

Sub YS_SumTotals()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("YearSpendatures")

    Dim months As Variant
    Dim col As Long

    months = Array("January", "February", "March", "April", "May", "June", _
                  "July", "August", "September", "October", "November", "December")

    ' Row 5 formulas from Table9
    For col = 0 To 11 ' D to O = 4 to 15
        ws.Cells(5, col + 4).FormulaR1C1 = "=SUBTOTAL(109,Table9[" & months(col) & "])"
    Next col

    ' P5 = sum of D5 to O5
    ws.Cells(5, 16).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"

    ' Row 25 formulas from Table10
    For col = 0 To 11 ' D to O = 4 to 15
        ws.Cells(25, col + 4).FormulaR1C1 = "=SUBTOTAL(109,Table10[" & months(col) & "])"
    Next col

    ' P25 = sum of D25 to O25
    ws.Cells(25, 16).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"

    ' D26:P26 = Row 5 - Row 25
    For col = 4 To 16
        ws.Cells(26, col).FormulaR1C1 = "=R[-21]C - R[-1]C"
    Next col

    ws.Range("A1").Select
End Sub

Sub Cell_Center()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    With ws.Range("D2:P5")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    With ws.Range("D7:P26")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    With ws.Range("E30:E119")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub RefreshExpensePivotTableSilent()
    ' Silently refreshes the ExpensePvt pivot table on the Expenses_pvt sheet
    On Error Resume Next
    
    ThisWorkbook.Sheets("Expenses_pvt").PivotTables("ExpensePvt").RefreshTable
    
    On Error GoTo 0
End Sub

