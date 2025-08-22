Attribute VB_Name = "Module4"
Sub ImprovedShoppingListMacro_Simplified()
    Dim ws As Worksheet, shoppingWs As Worksheet
    Dim lastRow As Long, purchaseCount As Long
    Dim i As Long, copyRow As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    
    Set ws = ActiveWorkbook.Worksheets("StorageData")
    Set shoppingWs = ActiveWorkbook.Worksheets("ShoppingList")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    
    ' Calculate QTY to Buy in col I
    For i = 2 To lastRow
        Dim currentQty As Double, preferredQty As Double
        currentQty = Val(ws.Cells(i, 3).Value)
        preferredQty = Val(ws.Cells(i, 7).Value)
        If currentQty < preferredQty Then
            ws.Cells(i, 8).Value = "Purchase Item"
            ws.Cells(i, 9).Value = preferredQty - currentQty
        Else
            ws.Cells(i, 8).Value = ""
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    ws.Range("I2:I" & lastRow).NumberFormat = "0"
    
    ' Sort: Status desc, Item asc
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("H2:H" & lastRow), Order:=xlDescending
        .SortFields.Add2 Key:=ws.Range("A2:A" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:I" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    purchaseCount = Application.WorksheetFunction.CountIf(ws.Range("H2:H" & lastRow), "Purchase Item")
    If purchaseCount = 0 Then
        MsgBox "No items need to be purchased at this time.", vbInformation
        GoTo CleanUp
    End If
    
    ' Clear and set headers
    shoppingWs.Cells.Clear
    shoppingWs.Range("A1").Value = "Item"
    shoppingWs.Range("B1").Value = "Storage QTY"
    shoppingWs.Range("C1").Value = "Preferred QTY"
    shoppingWs.Range("D1").Value = "QTY to Buy"
    shoppingWs.Range("A1:D1").Font.Bold = True
    shoppingWs.Range("A1:D1").Interior.ColorIndex = 15
    shoppingWs.Range("A1:D1").Borders.LineStyle = xlContinuous
    
    ' Copy only columns A, C, G, I
    copyRow = 2
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value = "Purchase Item" Then
            shoppingWs.Cells(copyRow, 1).Value = ws.Cells(i, 1).Value
            shoppingWs.Cells(copyRow, 2).Value = ws.Cells(i, 3).Value
            shoppingWs.Cells(copyRow, 3).Value = ws.Cells(i, 7).Value
            shoppingWs.Cells(copyRow, 4).Value = ws.Cells(i, 9).Value
            copyRow = copyRow + 1
        End If
    Next i
    
    ' Format numbers
    shoppingWs.Range("B2:D" & copyRow - 1).NumberFormat = "0"
    shoppingWs.Range("A1:D" & copyRow - 1).Borders.LineStyle = xlContinuous
    shoppingWs.Columns("A:D").AutoFit
    
    MsgBox "Shopping list created successfully! Found " & purchaseCount & " items.", vbInformation
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.description, vbCritical
    GoTo CleanUp
End Sub
