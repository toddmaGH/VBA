VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Donation_Frm 
   Caption         =   "Donations Form"
   ClientHeight    =   2892
   ClientLeft      =   228
   ClientTop       =   936
   ClientWidth     =   6192
   OleObjectBlob   =   "Donation_Frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Donation_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === Inside Donation_Frm code module ===
Option Explicit
Public ParentBudgetForm As BudgetFrm

Private Sub donationSubmit_cmd_Click()
' Set the form size first
    Me.Height = 148.8
    Me.Width = 268.8
    
    ' Optional: Center the form on screen
    Me.StartUpPosition = 0 ' Manual positioning
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    
On Error GoTo ErrorHandler
    Dim wsBudget As Worksheet
    Dim wsYearSpendatures As Worksheet
    Dim wsAggregate As Worksheet
    Dim budgetMonth As String
    Dim donationAmount As Double
    Dim description As String
    Dim targetRow As Long
    Dim i As Long
    
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    Set wsYearSpendatures = ThisWorkbook.Sheets("YearSpendatures")
    Set wsAggregate = ThisWorkbook.Sheets("Donations_Aggregate")
    
    ' Modified validation - now allows blank/empty values and zeros
    If Me.DonationFrm_txt.Value = "" Then
        ' Allow blank entries - set donation amount to 0
        donationAmount = 0
    ElseIf Not IsNumeric(Me.DonationFrm_txt.Value) Then
        ' Only reject if non-numeric and not blank
        MsgBox "Please enter a valid donation amount or leave blank for zero.", vbExclamation, "Invalid Input"
        Me.DonationFrm_txt.SetFocus
        Exit Sub
    Else
        ' Convert to numeric value (allows zeros and positive numbers)
        donationAmount = Val(Me.DonationFrm_txt.Value)
    End If
    
    ' Get description from ComboBox
    description = Me.Donation_cbo.Value
    
    ' Validate that a donation type has been selected
    If description = "" Then
        MsgBox "Please select a donation type.", vbExclamation, "Selection Required"
        Me.Donation_cbo.SetFocus
        Exit Sub
    End If
    
    ' Removed the validation that rejected zero amounts
    ' Now accepts any non-negative amount including zero
    If donationAmount < 0 Then
        MsgBox "Donation amount cannot be negative.", vbExclamation, "Invalid Amount"
        Me.DonationFrm_txt.SetFocus
        Exit Sub
    End If
    
    ' Update parent form field
    If Not ParentBudgetForm Is Nothing Then
        ParentBudgetForm.donations_txt.Value = donationAmount
    End If
    
    ' Get the budget month from A1 or fallback to combo box
    budgetMonth = Trim(wsBudget.Range("A1").Value)
    If budgetMonth = "" And Not ParentBudgetForm Is Nothing Then
        If ParentBudgetForm.month_cbo.ListIndex >= 0 Then
            budgetMonth = ParentBudgetForm.month_cbo.Value
            wsBudget.Range("A1").Value = budgetMonth
        Else
            MsgBox "Please select a month in the Budget form.", vbExclamation, "Month Required"
            Exit Sub
        End If
    ElseIf budgetMonth = "" Then
        MsgBox "No month available. Please select a month.", vbExclamation, "Month Required"
        Exit Sub
    End If
    
    ' Find the next blank cell in column B starting from row 29
    targetRow = 29
    Do While wsYearSpendatures.Cells(targetRow, "B").Value <> ""
        targetRow = targetRow + 1
    Loop
    
    ' Write the donation data
    wsYearSpendatures.Cells(targetRow, "B").Value = budgetMonth
    wsYearSpendatures.Cells(targetRow, "C").Value = Date
    wsYearSpendatures.Cells(targetRow, "D").Value = donationAmount
    wsYearSpendatures.Cells(targetRow, "E").Value = description
    
    ' Optional aggregate call
    On Error Resume Next
    Call UpdateDonationAggregate
    On Error GoTo ErrorHandler
    
    ' Updated success message to handle zero amounts appropriately
    If donationAmount = 0 Then
        MsgBox "Zero donation entry for " & budgetMonth & " saved!", vbInformation, "Donation Saved"
    Else
        MsgBox "Donation of " & Format(donationAmount, "$#,##0.00") & " for " & budgetMonth & " saved!", vbInformation, "Donation Saved"
    End If
    
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in donationSubmit_cmd_Click: " & Err.description, vbCritical, "Error"
End Sub

' This event fires when the form is initialized - this will populate the combo box
Private Sub UserForm_Initialize()
    Call PopulateDonationComboBox
End Sub

Private Sub PopulateDonationComboBox()
    Dim ws As Worksheet
    Dim organizationRange As Range
    Dim cell As Range
    Dim lastRow As Long
    
    ' Set the worksheet containing your item list
    Set ws = ThisWorkbook.Sheets("ItemList") ' Change to your actual sheet name
    
    ' Clear existing items
    Me.Donation_cbo.Clear
    
    ' Find the last row with data in column A (assuming items are in column A)
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Skip if only header row or no data
    If lastRow < 2 Then Exit Sub
    
    ' Set the range (starting from row 2 to skip header)
    Set organizationRange = ws.Range("H2:H" & lastRow)
    
    ' Add each item to the ComboBox (skip empty cells)
    For Each cell In organizationRange
        If Trim(cell.Value) <> "" Then
            Me.Donation_cbo.AddItem cell.Value
        End If
    Next cell
End Sub
