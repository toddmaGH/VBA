Attribute VB_Name = "Module5"
' ========================================================================
'
    ' Author: Todd Martin
    ' Date: 8/1/2025
    ' Dependencies: Requires worksheets "YearSpendatures" and "Budget"
    '
' ========================================================================
Option Explicit
Sub AddBudgetDropdownToToolbar()
    Dim bar As CommandBar
    Dim dropdown As CommandBarPopup
    Dim existingCtrl As CommandBarControl
    Dim btnBudget As CommandBarButton
    Dim btnDonation As CommandBarButton
    Dim btnDashboard As CommandBarButton
    
    ' Access the Worksheet Menu Bar
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    If bar Is Nothing Then
        MsgBox "The 'Worksheet Menu Bar' was not found.", vbCritical
        Exit Sub
    End If
    ' Remove existing "Budget Tools" dropdown if present
    For Each existingCtrl In bar.Controls
        If existingCtrl.caption = "Budget Tools" Then
            existingCtrl.Delete
            Exit For
        End If
    Next existingCtrl
    ' Add new dropdown menu
    Set dropdown = bar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With dropdown
        .caption = "Budget Tools"
        .Tag = "BudgetToolsDropdown"
    End With
    ' Add "Open Budget Form" item
    Set btnBudget = dropdown.Controls.Add(Type:=msoControlButton)
    With btnBudget
        .caption = "Open Budget Form"
        .onAction = "ShowBudgetForm"
        .faceId = 1594
    End With
    
    ' Add "Open hidden Dashboard" item
    Set btnDashboard = dropdown.Controls.Add(Type:=msoControlButton)
    With btnDashboard
        .caption = "View Dashboard"
        .onAction = "ShowDashboardSheet"
        .faceId = 984
    End With
    
    ' Add "Open hidden Donation Aggregate" item
    Set btnDashboard = dropdown.Controls.Add(Type:=msoControlButton)
    With btnDashboard
        .caption = "Donations Aggregate"
        .onAction = "CreateDonationsAggregateSilent"
        .faceId = 984
    End With
    
    ' Add "Open Storage Form" item
    Set btnDonation = dropdown.Controls.Add(Type:=msoControlButton)
    With btnDonation
        .caption = "Shopping/Storage"
        .onAction = "ShowStorage_Frm"
        .faceId = 1594
    End With
    
    ' Add "Show End Of Year Aggregate" item
    Set btnDonation = dropdown.Controls.Add(Type:=msoControlButton)
    With btnDonation
        .caption = "View EOY Aggregate"
        .onAction = "SumEOYTotals"
        .faceId = 1685
    End With
End Sub

' ============================================
' MISSING MACRO - ADD THIS TO YOUR MODULE
' ============================================

Sub ShowDonations_AggregateSheet()
    '
    ' Purpose: Show the Donations_Aggregate sheet
    '
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Check if the sheet exists
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Donations_Aggregate")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "The 'Donations_Aggregate' sheet was not found in this workbook.", vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Make the sheet visible if it's hidden
    If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVisible
    End If
    
    ' Activate the sheet
    ws.Activate
    
    ' Optional: Select cell A1
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing Donations_Aggregate sheet: " & Err.description, vbCritical, "Error"
    
End Sub

' ============================================
' OPTIONAL: ADD THESE OTHER MACROS IF MISSING
' ============================================
Sub ShowBudgetForm()
    '
    ' Purpose: Show the Budget Form
    '
    
    On Error GoTo ErrorHandler
    
    ' Show the Budget Form directly
    BudgetFrm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing Budget Form: " & Err.description, vbCritical, "Error"
    
End Sub

Sub ShowDashboardSheet()
    '
    ' Purpose: Show the Dashboard sheet
    '
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Check if the sheet exists
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "The 'Dashboard' sheet was not found in this workbook.", vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Make the sheet visible if it's hidden
    If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVisible
    End If
    
    ' Activate the sheet
    ws.Activate
    
    ' Optional: Select cell A1
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing Dashboard sheet: " & Err.description, vbCritical, "Error"
    
End Sub

Sub ShowStorage_Frm()
    '
    ' Purpose: Show the Storage Form
    '
    
    On Error GoTo ErrorHandler
    
    ' Replace "Storage_Frm" with your actual form name if different
    Load Storage_Frm
    Storage_Frm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing Storage Form: " & Err.description, vbCritical, "Error"
    
End Sub

' Helper procedure to open Donation_Frm
Sub ShowDonationForm()
    Donation_Frm.Show
End Sub

Sub RemoveBudgetDropdownFromToolbar()
    Dim ctrl As CommandBarControl
    Dim bar As CommandBar

    On Error Resume Next
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    If Not bar Is Nothing Then
        For Each ctrl In bar.Controls
            If ctrl.caption = "Budget Tools" Then
                ctrl.Delete
                Exit For
            End If
        Next ctrl
    End If
    On Error GoTo 0
End Sub
Sub ShoweoyaggSheet()
    '
    ' Purpose: Show the Dashboard sheet
    '
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Check if the sheet exists
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("EOY_Aggregate")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "The 'EOY_Aggregate' sheet was not found in this workbook.", vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Make the sheet visible if it's hidden
    If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVisible
    End If
    
    ' Activate the sheet
    ws.Activate
    
    ' Optional: Select cell A1
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing EOY_Aggregate sheet: " & Err.description, vbCritical, "Error"
    
End Sub
Sub ShowEOYAggregate()
    '
    ' Purpose: Show EOY_Aggregate sheet and execute related processing
    ' Author: Todd Martin
    ' Date: 8/10/2025
    '
    
    On Error GoTo ErrorHandler
    
    ' Make the EOY_Aggregate sheet visible and activate it
    With ThisWorkbook.Sheets("EOY_Aggregate")
        .Visible = xlSheetVisible
        .Activate
    End With
    
    ' Check if EOY_Agg subroutine exists before calling it
    If SubroutineExists("EOY_Agg") Then
        Call EOY_Agg
    Else
        MsgBox "The subroutine 'EOY_Agg' was not found. Please ensure it exists in this workbook.", _
               vbExclamation, "Subroutine Not Found"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description, vbCritical, "Error in ShowEOYAggregate"
    
    Call ShoweoyaggSheet
    
End Sub
Private Sub AddToolbarButton(dropdown As CommandBarPopup, caption As String, onAction As String, faceId As Long)
    ' Helper to add consistent toolbar buttons
    Dim btn As CommandBarButton
    Set btn = dropdown.Controls.Add(Type:=msoControlButton)
    With btn
        .caption = caption
        .onAction = onAction
        .faceId = faceId
    End With
End Sub

' ========================================================================
' UTILITY FUNCTIONS
' ========================================================================

Public Function IsFormLoaded(formName As String) As Boolean
    ' Enhanced form checking
    Dim i As Integer
    For i = 0 To UserForms.Count - 1
        If StrComp(UserForms(i).Name, formName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    IsFormLoaded = False
End Function

' Simplified alias for backward compatibility
Function IsLoaded(formName As String) As Boolean
    IsLoaded = IsFormLoaded(formName)
End Function
Public Sub UpdateDonationAggregate()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim dict As Object
    Dim i As Long
    Dim monthName As String
    Dim amount As Double
    Dim ws As Worksheet
    Dim formulaRange As Range
    Dim budgetWS As Worksheet
    Dim monthValue As String
    Dim lastRow As Long
    Dim totalCell As Range
    
    Set ws = ThisWorkbook.Sheets("Donations_Aggregate")
    Set budgetWS = ThisWorkbook.Sheets("Budget")
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    Set wsDest = ThisWorkbook.Sheets("Donations_Aggregate")
    Set dict = CreateObject("Scripting.Dictionary")
    
    monthValue = Trim(budgetWS.Range("A1").Value)
    If Len(monthValue) = 0 Then Exit Sub
    
    ' Ensure Column F has exactly 17 rows (F2:F18) - Fixed data that shouldn't be deleted
    ' Only populate if F2:F18 is empty or needs updating
    If wsDest.Range("F2").Value = "" Then
        ' Add your fixed category data here - this should be permanent data
        wsDest.Range("F2").Value = "Category 1"
        wsDest.Range("F3").Value = "Category 2"
        wsDest.Range("F4").Value = "Category 3"
        wsDest.Range("F5").Value = "Category 4"
        wsDest.Range("F6").Value = "Category 5"
        wsDest.Range("F7").Value = "Category 6"
        wsDest.Range("F8").Value = "Category 7"
        wsDest.Range("F9").Value = "Category 8"
        wsDest.Range("F10").Value = "Category 9"
        wsDest.Range("F11").Value = "Category 10"
        wsDest.Range("F12").Value = "Category 11"
        wsDest.Range("F13").Value = "Category 12"
        wsDest.Range("F14").Value = "Category 13"
        wsDest.Range("F15").Value = "Category 14"
        wsDest.Range("F16").Value = "Category 15"
        wsDest.Range("F17").Value = "Category 16"
        wsDest.Range("F18").Value = "Category 17"
    End If
    
    ' Set formula range for G2:G17 (17 rows)
    Set formulaRange = ws.Range("G2:G17")
    
    ' CORRECTED FORMULA: Lookup Donations_Aggregate F column in YearSpendatures E column
    ' and sum corresponding values from YearSpendatures D column
    formulaRange.FormulaR1C1 = _
        "=SUMIF(YearSpendatures!R30C5:R100C5,Donations_Aggregate!RC[-1],YearSpendatures!R30C4:R100C4)"
    
    ' Calculate sum of G2:G17 and place result in G18
    ws.Range("G18").FormulaR1C1 = "=SUM(R[-16]C:R[-1]C)"  ' G2:G17 = 16 rows (from current row -16 to -1)
    
    ' Convert all formulas to values in column G
    Dim gRange As Range
    Set gRange = ws.Range("G2:G18")
    gRange.Value = gRange.Value  ' This converts formulas to values
    
    ' Rest of your existing logic for dictionary population
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
    
    ' Remove any existing Total row before appending
    Set totalCell = wsDest.Columns("A").Find(What:="Total", LookIn:=xlValues, LookAt:=xlWhole)
    If Not totalCell Is Nothing Then wsDest.Rows(totalCell.Row).Delete
    
    ' Add current month if it doesn't exist
    If wsDest.Columns("A").Find(What:=monthValue, LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
        lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
        wsDest.Cells(lastRow, "A").Value = monthValue
        If dict.Exists(monthValue) Then
            wsDest.Cells(lastRow, "B").Value = dict(monthValue)
        Else
            wsDest.Cells(lastRow, "B").Value = 0
        End If
    End If
    
    ' Add Total row
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
    wsDest.Cells(lastRow, "A").Value = "Total"
    wsDest.Cells(lastRow, "B").Formula = "=SUM(B2:B" & lastRow - 1 & ")"
    
    ' Convert the Total formula to value
    wsDest.Cells(lastRow, "B").Value = wsDest.Cells(lastRow, "B").Value
    
    ' Format the Total row
    wsDest.Range("A1:B1").Copy
    wsDest.Range("A" & lastRow & ":B" & lastRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' Apply number formatting
    wsDest.Range("B2:B" & lastRow).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    
    ' Convert all data to values only (no formulas) - Final step to ensure everything is values
    Dim dataRange As Range
    Set dataRange = wsDest.UsedRange
    dataRange.Value = dataRange.Value
    
End Sub
' Alternative version if you want to create a placeholder EOY_Agg subroutine
Sub ShowEOYAggregateWithPlaceholder()
    '
    ' Purpose: Show EOY_Aggregate sheet and execute related processing
    ' This version includes error handling for missing subroutine
    '
    
    On Error GoTo ErrorHandler
    
    ' Make the EOY_Aggregate sheet visible and activate it
    With ThisWorkbook.Sheets("EOY_Aggregate")
        .Visible = xlSheetVisible
        .Activate
    End With
    
    ' Call the EOY processing subroutine
    Call EOY_Agg_Safe
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description, vbCritical, "Error in ShowEOYAggregate"
    
End Sub

' === TROUBLESHOOTING SOLUTIONS ===

'''' SOLUTION 1: If EOY_Agg exists but has a different name, rename the call:
'''Sub ShowEOYAggregate_Solution1()
'''    ThisWorkbook.Sheets("EOY_Aggregate").Visible = xlSheetVisible
'''    ThisWorkbook.Sheets("EOY_Aggregate").Activate
'''
'''    ' Try these alternative names (uncomment the correct one):
'''    ' Call EOYAgg
'''    ' Call EOY_Aggregate
'''    ' Call ProcessEOYData
'''    ' Call CalculateEOYAggregates
'''
'''End Sub

' SOLUTION 2: If you need to create the EOY_Agg subroutine from scratch:
Sub EOY_Agg()
    '
    ' Purpose: Process End of Year aggregate data
    ' Add your specific EOY processing logic here
    '
    
    ' Example processing (replace with your actual needs):
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("EOY_Aggregate")
    
    ' Your EOY processing code goes here
    'MsgBox "EOY_Agg subroutine executed successfully!", vbInformation
    
    Call ShowEOYAggregate_WithModule
    
    Set ws = Nothing
End Sub

' SOLUTION 3: If the subroutine is in a different module, specify the module:
Sub ShowEOYAggregate_WithModule()
    ThisWorkbook.Sheets("EOY_Aggregate").Visible = xlSheetVisible
    ThisWorkbook.Sheets("EOY_Aggregate").Activate
    
    ' If EOY_Agg is in a specific module, call it like this:
    ' Call ModuleName.EOY_Agg
    ' Example: Call Module1.EOY_Agg
    
End Sub


Sub CreateDonationsAggregateSilent()
    ' Complete Donations Aggregate Processing
    ' Part 1: Monthly Aggregation (Columns A & B)
    ' Part 2: Organization Aggregation (Columns F & G)
    
    On Error GoTo ErrorHandler
    
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim i As Long
    Dim orgName As String, monthName As String
    Dim orgAmount As Variant, monthAmount As Variant
    Dim orgDict As Object, monthDict As Object
    Dim orgKey As String, monthKey As String
    Dim targetRow As Long
    Dim totalAmount As Double, monthTotalAmount As Double
    Dim totalRow As Long
    Dim dictKey As Variant
    
    ' Step 1: Open (activate) the Donations_Aggregate sheet
    Call ShowDonationAggregate
    
    Set wsTarget = ThisWorkbook.Sheets("Donations_Aggregate")
    wsTarget.Activate
    
    ' Set source worksheet reference
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    
    ' Create Dictionary objects for better handling of unique keys
    Set orgDict = CreateObject("Scripting.Dictionary")
    Set monthDict = CreateObject("Scripting.Dictionary")
    orgDict.CompareMode = 1 ' Text compare, case-insensitive
    monthDict.CompareMode = 1 ' Text compare, case-insensitive
    
    ' Find last row of data in source sheet starting from row 30
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRowSource < 30 Then lastRowSource = 30 ' Ensure we start at row 30
    
    ' PART 1: MONTHLY AGGREGATION (Columns A & B)
    ' Step 2: Clear columns A1:B14 completely
    wsTarget.Range("A1:B14").ClearContents
    
    ' Step 3: Set headers for monthly data
    wsTarget.Range("A1").Value = "Donation Month"
    wsTarget.Range("B1").Value = "Donations"
    wsTarget.Range("A1:B1").Font.Bold = True
    
    ' Step 4: Process unique months from YearSpendatures B30:B
    ' Step 5: Sum amounts for each month from column D30:D
    For i = 30 To lastRowSource
        monthName = Trim(CStr(wsSource.Cells(i, 2).Value)) ' Column B - Month name
        monthAmount = wsSource.Cells(i, 4).Value ' Column D - Amount
        
        ' Only process if month name is not empty and amount is numeric
        If monthName <> "" And IsNumeric(monthAmount) Then
            monthKey = UCase(monthName) ' Use uppercase for case-insensitive comparison
            
            ' Check if month already exists in dictionary
            If monthDict.Exists(monthKey) Then
                ' Month exists, add to existing total
                monthDict(monthKey) = monthDict(monthKey) + CDbl(monthAmount)
            Else
                ' Month not found, add new entry
                monthDict.Add monthKey, CDbl(monthAmount)
            End If
        End If
    Next i
    
    ' Step 6: Paste monthly data starting at A2:B2
    ' Define standard month order for consistent display
    Dim monthOrder As Variant
    monthOrder = Array("JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", _
                      "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER")
    
    targetRow = 2
    monthTotalAmount = 0 ' Initialize total for calculation
    
    ' Loop through months in proper order (A2:A13)
    For i = 0 To 11 ' 12 months maximum
        If targetRow > 13 Then Exit For ' Stop at row 13
        
        monthKey = monthOrder(i)
        If monthDict.Exists(monthKey) Then
            ' Month has data, paste it
            wsTarget.Cells(targetRow, 1).Value = StrConv(monthKey, vbProperCase) ' Column A
            wsTarget.Cells(targetRow, 2).Value = monthDict(monthKey) ' Column B
            monthTotalAmount = monthTotalAmount + monthDict(monthKey)
        Else
            ' Month has no data, show 0
            wsTarget.Cells(targetRow, 1).Value = StrConv(monthKey, vbProperCase) ' Column A
            wsTarget.Cells(targetRow, 2).Value = 0 ' Column B
        End If
        targetRow = targetRow + 1
    Next i
    
    ' Step 7: Add Total in row 14
    wsTarget.Range("A14").Value = "Total"
    wsTarget.Range("B14").Value = monthTotalAmount
    wsTarget.Range("A14:B14").Font.Bold = True
    
    ' Format monthly amounts as currency
    wsTarget.Range("B2:B14").NumberFormat = "$#,##0.00"
    
    ' PART 2: ORGANIZATION AGGREGATION (Columns F & G)
    ' Clear columns F:G completely
    wsTarget.Columns("F:G").ClearContents
    
    ' Set headers for organization data
    wsTarget.Range("F1").Value = "Donated Organization"
    wsTarget.Range("G1").Value = "Donated Amount"
    wsTarget.Range("F1:G1").Font.Bold = True
    
    ' Process unique organizations from YearSpendatures E30:E and sum amounts from D30:D
    For i = 30 To lastRowSource
        orgName = Trim(CStr(wsSource.Cells(i, 5).Value)) ' Column E - Organization name
        orgAmount = wsSource.Cells(i, 4).Value ' Column D - Amount
        
        ' Only process if organization name is not empty and amount is numeric
        If orgName <> "" And IsNumeric(orgAmount) Then
            orgKey = UCase(orgName) ' Use uppercase for case-insensitive comparison
            
            ' Check if organization already exists in dictionary
            If orgDict.Exists(orgKey) Then
                ' Organization exists, add to existing total
                orgDict(orgKey) = orgDict(orgKey) + CDbl(orgAmount)
            Else
                ' Organization not found, add new entry
                orgDict.Add orgKey, CDbl(orgAmount)
            End If
        End If
    Next i
    
    ' Paste unique organizations and their summed amounts starting at F2:G2
    targetRow = 2
    totalAmount = 0 ' Initialize total for calculation
    
    ' Loop through dictionary and paste organization data
    For Each dictKey In orgDict.Keys
        ' Extract original organization name (convert back from uppercase key)
        orgName = StrConv(dictKey, vbProperCase) ' Convert to proper case
        
        ' Paste organization name to column F
        wsTarget.Cells(targetRow, 6).Value = orgName
        
        ' Paste summed amount to column G
        wsTarget.Cells(targetRow, 7).Value = orgDict(dictKey)
        
        ' Add to running total
        totalAmount = totalAmount + orgDict(dictKey)
        
        targetRow = targetRow + 1
    Next dictKey
    
    ' Add calculated "Total" row at the end of organization data
    If orgDict.Count > 0 Then
        totalRow = orgDict.Count + 2 ' Row after last organization entry
        
        ' Add "Total" label in column F
        wsTarget.Cells(totalRow, 6).Value = "Total"
        
        ' Add calculated sum in column G
        wsTarget.Cells(totalRow, 7).Value = totalAmount
        
        ' Format the total row
        wsTarget.Range("F" & totalRow & ":G" & totalRow).Font.Bold = True
        wsTarget.Cells(totalRow, 7).NumberFormat = "$#,##0.00"
    End If
    
    ' Format all organization amounts as currency
    If orgDict.Count > 0 Then
        wsTarget.Range("G2:G" & (orgDict.Count + 1)).NumberFormat = "$#,##0.00"
    End If
    
    ' Auto-fit columns for better display
    wsTarget.Columns("A:B").AutoFit
    wsTarget.Columns("F:G").AutoFit
    
    ' Clean up objects
    Set orgDict = Nothing
    Set monthDict = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    
    Exit Sub
    
    
    
ErrorHandler:
    MsgBox "Error in CreateDonationsAggregateSilent: " & Err.description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Processing row: " & i, vbCritical, "Donation Aggregate Error"
    
    ' Clean up objects on error
    Set orgDict = Nothing
    Set monthDict = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
End Sub


Sub ShowDonationAggregate()
    '
    ' Purpose: Show Donations Aggregate sheet and execute related processing
    ' Author: Todd Martin
    ' Date: 8/15/2025
    '
    
''''    On Error GoTo ErrorHandler
    
    ' Make the EOY_Aggregate sheet visible and activate it
    With ThisWorkbook.Sheets("Donations_Aggregate")
        .Visible = xlSheetVisible
        .Activate
    End With
    
        
''ErrorHandler:
''    MsgBox "An error occurred: " & Err.description, vbCritical, "Error in ShowDonationsAggregate"
End Sub

Sub ShowShoppingList()
    '
    ' Purpose: Show Shopping List execute related processing
    ' Author: Todd Martin
    ' Date: 8/18/2025
    '
    
''''    On Error GoTo ErrorHandler
    
    ' Make the EOY_Aggregate sheet visible and activate it
    With ThisWorkbook.Sheets("ShoppingList")
        .Visible = xlSheetVisible
        .Activate
    End With
    
        
''ErrorHandler:
''    MsgBox "An error occurred: " & Err.description, vbCritical, "Error in ShowDonationsAggregate"
End Sub
Sub ShowStorageData()
    '
    ' Purpose: Show StorageData sheet
    ' Author: Todd Martin
    ' Date: 8/18/2025
    '
    
''''    On Error GoTo ErrorHandler
    
    ' Make the EOY_Aggregate sheet visible and activate it
    With ThisWorkbook.Sheets("StorageData")
        .Visible = xlSheetVisible
        .Activate
    End With
    
        
''ErrorHandler:
''    MsgBox "An error occurred: " & Err.description, vbCritical, "Error in ShowDonationsAggregate"
End Sub
