Attribute VB_Name = "Module6"
Sub HighlightAndExportExpenseOutliers()
    ' Requires macro-enabled workbook to run
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, wsNew As Worksheet
    Dim r As Long, c As Long, lastRow As Long, lastCol As Long
    Dim skipRows As Variant, isSkip As Boolean
    Dim months As Variant, monthIdx As Long
    Dim arrValues() As Double, arrRowIdx() As Long, arrColIdx() As Long
    Dim i As Long, j As Long, n As Long
    Dim q1 As Double, q3 As Double, iqr As Double
    Dim outlierColor As Long
    Dim outlierCount As Long
    Dim sheetExists As Boolean
    Dim wsForm As Worksheet
    Dim iqrMultiplier As Double
    
    ' Check if source worksheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("YearSpendatures")
    Set wsForm = ThisWorkbook.Sheets("form controls")
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        MsgBox "Worksheet 'YearSpendatures' not found!", vbCritical
        Exit Sub
    End If
    If wsForm Is Nothing Then
        MsgBox "Worksheet 'form controls' not found!", vbCritical
        Exit Sub
    End If
    
    ' Get IQR multiplier from user input
    Dim userInput As String
    Dim currentValue As Double
    
    ' Check current value in form controls sheet for default
    On Error Resume Next
    currentValue = wsForm.Range("B2").Value
    If currentValue <= 0 Or currentValue > 5 Then currentValue = 1.5 ' Default if invalid
    On Error GoTo ErrorHandler
    
    ' Show input box with current/default value
    userInput = InputBox( _
        "What IQR Multiplier do you want to use?" & vbCrLf & vbCrLf & _
        "1.0 = High sensitivity (more outliers)" & vbCrLf & _
        "1.5 = Standard (recommended)" & vbCrLf & _
        "2.0 = Low sensitivity (fewer outliers)" & vbCrLf & vbCrLf & _
        "Enter value between 0.1 and 5.0:", _
        "IQR Multiplier Selection", _
        currentValue)
    
    ' Handle user cancellation or empty input
    If userInput = "" Then
        MsgBox "Analysis cancelled by user.", vbInformation
        Exit Sub
    End If
    
    ' Validate and convert input
    If Not IsNumeric(userInput) Then
        MsgBox "Invalid input. Please enter a numeric value.", vbCritical
        Exit Sub
    End If
    
    iqrMultiplier = CDbl(userInput)
    
    ' Validate IQR multiplier range
    If iqrMultiplier <= 0 Or iqrMultiplier > 5 Then
        MsgBox "Please enter a valid IQR multiplier between 0.1 and 5.0." & vbCrLf & _
               "You entered: " & iqrMultiplier, vbExclamation
        Exit Sub
    End If
    
    ' Save the value to form controls sheet
    wsForm.Range("B2").Value = iqrMultiplier
    
    ' Check if "Expense Outliers" sheet already exists and delete it
    sheetExists = False
    For Each wsNew In ThisWorkbook.Sheets
        If wsNew.Name = "Expense Outliers" Then
            Application.DisplayAlerts = False
            wsNew.Delete
            Application.DisplayAlerts = True
            sheetExists = True
            Exit For
        End If
    Next wsNew
    
    ' Create new worksheet
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ws)
    wsNew.Name = "Expense Outliers"
    
    ' Set up headers with formatting
    With wsNew.Range("A1:D1")
        .Value = Array("Item", "Month", "Value", "Outlier Analysis")
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
    End With
    
    skipRows = Array(7, 8, 10, 14, 19) ' Table10 rows to skip (relative to B6, so B7 = row 7, etc.)
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    outlierColor = RGB(255, 199, 206) ' Light red
    outlierCount = 0
    
    ' Collect all values for each month (excluding skipRows)
    For monthIdx = 1 To 12
        ' Build array of values for this month
        n = 0
        ReDim arrValues(1 To 100)
        ReDim arrRowIdx(1 To 100)
        
        For r = 7 To 24 ' Table10 data rows (B7:B24)
            isSkip = False
            For i = LBound(skipRows) To UBound(skipRows)
                If r = skipRows(i) Then
                    isSkip = True
                    Exit For
                End If
            Next i
            
            If Not isSkip Then
                ' Check if cell has numeric value
                If IsNumeric(ws.Cells(r, monthIdx + 3).Value) And ws.Cells(r, monthIdx + 3).Value <> "" Then
                    n = n + 1
                    If n > UBound(arrValues) Then
                        ReDim Preserve arrValues(1 To n + 50)
                        ReDim Preserve arrRowIdx(1 To n + 50)
                    End If
                    arrValues(n) = CDbl(ws.Cells(r, monthIdx + 3).Value) ' D=4, so Jan=4, Feb=5, etc.
                    arrRowIdx(n) = r
                End If
            End If
        Next r
        
        If n = 0 Then GoTo NextMonth
        
        ' Resize arrays to actual size
        ReDim Preserve arrValues(1 To n)
        ReDim Preserve arrRowIdx(1 To n)
        
        ' Only calculate quartiles if we have enough data points
        If n >= 4 Then
            ' Calculate Q1, Q3, IQR
            q1 = WorksheetFunction.Quartile(arrValues, 1)
            q3 = WorksheetFunction.Quartile(arrValues, 3)
            iqr = q3 - q1
            
            ' Check for outliers and highlight/copy
            For i = 1 To n
                If arrValues(i) < (q1 - iqrMultiplier * iqr) Or arrValues(i) > (q3 + iqrMultiplier * iqr) Then
                    ' Highlight outlier in original sheet
                    ws.Cells(arrRowIdx(i), monthIdx + 3).Interior.Color = outlierColor
                    
                    ' Copy to outlier sheet with analysis
                    outlierCount = outlierCount + 1
                    wsNew.Cells(outlierCount + 1, 1).Value = ws.Cells(arrRowIdx(i), 2).Value ' Item name (column B)
                    wsNew.Cells(outlierCount + 1, 2).Value = months(monthIdx - 1)
                    wsNew.Cells(outlierCount + 1, 3).Value = arrValues(i)
                    
                    ' Add outlier analysis in column D
                    Dim analysisText As String
                    Dim lowerBound As Double, upperBound As Double
                    lowerBound = q1 - iqrMultiplier * iqr
                    upperBound = q3 + iqrMultiplier * iqr
                    
                    If arrValues(i) < lowerBound Then
                        analysisText = "LOW: $" & Format(arrValues(i), "#,##0.00") & " is below expected range (Min: $" & Format(lowerBound, "#,##0.00") & "). Review for potential savings or data entry error."
                    Else
                        analysisText = "HIGH: $" & Format(arrValues(i), "#,##0.00") & " exceeds typical spending (Max: $" & Format(upperBound, "#,##0.00") & "). Investigate unusual expense or budget variance."
                    End If
                    
                    wsNew.Cells(outlierCount + 1, 4).Value = analysisText
                End If
            Next i
        End If
        
NextMonth:
    Next monthIdx
    
        With wsForm
            .Cells(1, 1).Value = "IQR Multiplier:"
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 2).Value = "Last Used Value: " & iqrMultiplier
            .Cells(1, 2).Font.Italic = True
            .Cells(1, 2).Font.Color = RGB(128, 128, 128)
            .Cells(3, 1).Value = "Value saved from last analysis:"
            .Cells(3, 1).Font.Bold = True
            .Cells(4, 1).Value = "Sensitivity Guide:"
            .Cells(4, 1).Font.Bold = True
            .Cells(5, 1).Value = "1.0 = High sensitivity (more outliers)"
            .Cells(6, 1).Value = "1.5 = Standard (recommended)"
            .Cells(7, 1).Value = "2.0 = Low sensitivity (fewer outliers)"
            .Columns("A:B").AutoFit
        End With
        
        With wsNew
            .Columns("A:D").AutoFit
        .Range("C:C").NumberFormat = "$#,##0.00"
        .Range("D:D").WrapText = True
        .Range("D:D").VerticalAlignment = xlTop
        If outlierCount > 0 Then
            .Range("A1:D" & outlierCount + 1).Borders.LineStyle = xlContinuous
            ' Set row height for better readability of analysis text
            .Range("2:" & outlierCount + 1).RowHeight = 45
        End If
    End With
    
    ' Add methodology explanation at the bottom
    If outlierCount > 0 Then
        Dim startRow As Long
        startRow = outlierCount + 3
        
        With wsNew
            .Cells(startRow, 1).Value = "METHODOLOGY:"
            .Cells(startRow, 1).Font.Bold = True
            .Cells(startRow + 1, 1).Value = "• Outliers identified using Interquartile Range (IQR) method with " & iqrMultiplier & "× multiplier"
            .Cells(startRow + 2, 1).Value = "• Values below Q1 - " & iqrMultiplier & "×IQR or above Q3 + " & iqrMultiplier & "×IQR are flagged as outliers"
            .Cells(startRow + 3, 1).Value = "• LOW outliers may indicate savings opportunities or data entry errors"
            .Cells(startRow + 4, 1).Value = "• HIGH outliers suggest unusual expenses requiring investigation"
            .Cells(startRow + 5, 1).Value = "• Review these items for budget accuracy and spending patterns"
            .Cells(startRow + 6, 1).Value = "• Adjust IQR multiplier in 'form controls' sheet for different sensitivity levels"
            
            .Range(.Cells(startRow + 1, 1), .Cells(startRow + 6, 4)).Merge
            .Cells(startRow + 1, 1).WrapText = True
            .Cells(startRow + 1, 1).VerticalAlignment = xlTop
        End With
    End If
    
    ' Display results
    If outlierCount > 0 Then
        MsgBox "Outlier analysis complete using IQR multiplier of " & iqrMultiplier & ". Found " & outlierCount & " outliers." & vbCrLf & _
               "Results copied to '" & wsNew.Name & "' sheet." & vbCrLf & _
               "Value saved to 'form controls' sheet for future use.", vbInformation
        wsNew.Activate
    Else
        MsgBox "No outliers found using IQR multiplier of " & iqrMultiplier & "." & vbCrLf & _
               "Try using a lower multiplier for higher sensitivity." & vbCrLf & _
               "Value saved to 'form controls' sheet for future use.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description, vbCritical
    If Not wsNew Is Nothing Then
        Application.DisplayAlerts = False
        wsNew.Delete
        Application.DisplayAlerts = True
    End If
End Sub

