Attribute VB_Name = "Module1"
' ========================================================================
'
    ' Author: Todd Martin
    ' Date: 7/26/2025
    ' Dependencies: Requires worksheets "YearSpendatures" and "Budget"
    '
' ========================================================================
Option Explicit


Private Sub ProcessBudgetMonthData()
    ' Separated logic for budget month processing
    Dim wsBudget As Worksheet
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    
    Dim monthName As String, i As Long, matchRow As Long
    monthName = wsBudget.Range("A1").Value
    
    ' Search L3:L14 for monthName and update corresponding M and N columns
    For i = 3 To 14
        If StrComp(Trim(wsBudget.Cells(i, "L").Value), Trim(monthName), vbTextCompare) = 0 Then
            matchRow = i + 3
            wsBudget.Cells(matchRow, "M").Value = wsBudget.Range("E6").Value
            wsBudget.Cells(matchRow, "N").Value = wsBudget.Range("E31").Value
            Exit For
        End If
    Next i
    
    ' Calculate O6:O14 = M - N and update totals
    For i = 6 To 14
        wsBudget.Cells(i, "O").Value = wsBudget.Cells(i, "M").Value - wsBudget.Cells(i, "N").Value
    Next i
    
    ' Update sum formulas
    With wsBudget
        .Range("M15").Value = Application.Sum(.Range("M6:M14"))
        .Range("N15").Value = Application.Sum(.Range("N6:N14"))
        .Range("M16").Value = .Range("M15").Value + .Range("N15").Value
        .Range("O15").Value = Application.Sum(.Range("O6:O14"))
    End With
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
' ========================================================================
' DATA TRANSFER OPERATIONS
' ========================================================================

Private Sub ConvertFormulasToValues(targetRange As Range)
    '
    ' Purpose: Convert formulas in a range to their calculated values
    ' Parameters: targetRange - The range containing formulas to convert
    '
    
    On Error Resume Next
    
    With targetRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                      SkipBlanks:=False, Transpose:=False
    End With
    
    Application.CutCopyMode = False
    
End Sub
''' The following is a placeholder for the corrected ProcessMonthlyDonations sub.
Private Sub ProcessMonthlyDonations(wsSource As Worksheet, wsDest As Worksheet, monthValue As String)
    ' Separated donation processing logic for better organization
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, monthName As String, amount As Double
    
    ' Build dictionary of monthly totals
    For i = 30 To 200
        monthName = Trim(wsSource.Cells(i, "B").Value)
        amount = Val(wsSource.Cells(i, "D").Value)
        If Len(monthName) > 0 And amount <> 0 Then
            If dict.Exists(monthName) Then
                dict(monthName) = dict(monthName) + amount
            Else
                dict.Add monthName, amount
            End If
        End If
    Next i
    
    ' Clean up existing Total row
    Dim totalCell As Range
    Set totalCell = wsDest.Columns("A").Find(What:="Total", LookIn:=xlValues, LookAt:=xlWhole)
    If Not totalCell Is Nothing Then wsDest.Rows(totalCell.Row).Delete
    
    ' Add current month if missing
    If wsDest.Columns("A").Find(What:=monthValue, LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
        Dim lastRow As Long
        lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
        wsDest.Cells(lastRow, "A").Value = monthValue
        wsDest.Cells(lastRow, "B").Value = IIf(dict.Exists(monthValue), dict(monthValue), 0)
    End If
    
    ' Add formatted Total row
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
    wsDest.Cells(lastRow, "A").Value = "Total"
    wsDest.Cells(lastRow, "B").Formula = "=SUM(B2:B" & lastRow - 1 & ")"
    
    ' Apply formatting
    wsDest.Range("A1:B1").Copy
    wsDest.Range("A" & lastRow & ":B" & lastRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    wsDest.Range("B2:B" & lastRow).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
End Sub

' Helper function to check if a form is loaded
Public Function IsFormLoaded(formName As String) As Boolean
    Dim i As Integer
    For i = 0 To UserForms.Count - 1
        If UserForms(i).Name = formName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    IsFormLoaded = False
End Function

' Helper: Find table by name across all worksheets
Private Function FindTableByName(tableName As String) As ListObject
    Dim sh As Worksheet, lo As ListObject
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = sh.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set FindTableByName = lo
            Exit Function
        End If
    Next sh
End Function

' Helper: Get table column range address
Private Function TableColumnRangeAddress(lo As ListObject, colName As String) As String
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    If Not lc Is Nothing And Not lc.DataBodyRange Is Nothing Then
        TableColumnRangeAddress = lc.DataBodyRange.Address(External:=True)
    Else
        TableColumnRangeAddress = ""
    End If
End Function


