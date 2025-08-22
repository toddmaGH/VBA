Attribute VB_Name = "Module6"
'''## This is the code we want ##

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
    Set formulaRange = ws.Range("G2:G18")  ' Updated to G2:G18 as specified
    Set wsSource = ThisWorkbook.Sheets("YearSpendatures")
    Set wsDest = ThisWorkbook.Sheets("Donations_Aggregate")
    Set dict = CreateObject("Scripting.Dictionary")
    
    monthValue = Trim(budgetWS.Range("A1").Value)
    If Len(monthValue) = 0 Then Exit Sub
    
    ' CORRECTED FORMULA: Lookup Donations_Aggregate F column in YearSpendatures E column
    ' and sum corresponding values from YearSpendatures D column
    formulaRange.FormulaR1C1 = _
        "=SUMIF(YearSpendatures!R30C5:R100C5,Donations_Aggregate!RC[-1],YearSpendatures!R30C4:R100C4)"
    
    ' Update the total formula to match the new range
    ws.Range("G19").FormulaR1C1 = "=SUM(R[-17]C:R[-1]C)"  ' G2:G18 = 17 rows
    
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
    
    ' Format the Total row
    wsDest.Range("A1:B1").Copy
    wsDest.Range("A" & lastRow & ":B" & lastRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' Apply number formatting
    wsDest.Range("B2:B" & lastRow).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    
    ' Call other procedures
'''    Call Agg_Color
'''    Call StaticDonation
'''    Call BoldTotalCells
End Sub

Sub StaticDonation()
'
' StaticDonation Macro

    Range("A2:G92").Select
    Selection.Copy
    Selection.End(xlUp).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
End Sub

