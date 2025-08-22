VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Storage_Frm 
   Caption         =   "Storage"
   ClientHeight    =   7644
   ClientLeft      =   228
   ClientTop       =   996
   ClientWidth     =   16260
   OleObjectBlob   =   "Storage_Frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Storage_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form-level variables to store current quantities
Dim currentMeatQty As Double
Dim currentVegiQty As Double
Dim currentFruitQty As Double
Dim currentDriedQty As Double
Dim currentSoupsQty As Double
Dim currentPastaQty As Double
Dim currentLegumesQty As Double
Dim currentGrainsQty As Double
Dim currentOilFatsQty As Double
Dim currentDairyAltQty As Double
Dim currentBakingQty As Double
Dim currentSnacksQty As Double
Dim currentBevQty As Double
Dim currentFrozenQty As Double

' Exit button functionality
Private Sub cmdExit_Click()
    ' Unloads the UserForm, closing it without saving any changes
    Unload Me
End Sub

Private Sub Storage_Frm_Initialize()
    ' Initializes the UserForm, populates combo boxes, resets displays, and checks for low inventory
    On Error GoTo ErrorHandler

    ' Initialize Me.Tag
    Me.Tag = ""

    ' First populate combo boxes and initialize form
    Call PopulateItemComboBox
    Call InitializeCategoryLabels
    Call InitializeToBuyLabels
    
    ' Initialize quantity labels to default value
    Me.meatqtylbl.caption = "0"
    Me.vegiqtylbl.caption = "0"
    Me.fruitqtylbl.caption = "0"
    Me.driedfoodqtylbl.caption = "0"
    Me.soupqtylbl.caption = "0"
    Me.pastyqtylbl.caption = "0"
    Me.legumesqtylbl.caption = "0"
    Me.grainsqtylbl.caption = "0"
    Me.oilqtylbl.caption = "0"
    Me.dairtyqtylbl.caption = "0"
    Me.bakingqtylbl.caption = "0"
    Me.tratsqtylbl.caption = "0"
    Me.beveragesqtylbl.caption = "0"
    Me.Treatsqtylbl.caption = "0"

    ' Set the form size and position
    Me.Height = 410.4
    Me.Width = 823.8
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2

    ' Now check for low inventory after form is set up
    Call CheckLowInventoryAndPrompt

    Exit Sub
    
    Call ImprovedShoppingListMacro
    
ErrorHandler:
    MsgBox "Error initializing Storage Form: " & Err.description, vbCritical, "Initialization Error"
End Sub

Private Sub InitializeToBuyLabels()
    ' Sets default captions for "To Buy" labels on the form
    On Error Resume Next

    Me.canmeatstobuy_lbl.caption = "0"
    Me.cannedvegtobuy_lbl.caption = "0"
    Me.cannedfruittobuy_lbl.caption = "0"
    Me.driedfoodtobuy_lbl.caption = "0"
    Me.soupstobuy_lbl.caption = "0"
    Me.pastatobuy_lbl.caption = "0"
    Me.legumestobuy_lbl.caption = "0"
    Me.grainstobuy_lbl.caption = "0"
    Me.oilstobuy_lbl.caption = "0"
    Me.dairytobuy_lbl.caption = "0"
    Me.bakingtobut_lbl.caption = "0"
    Me.snacktobuy_lbl.caption = "0"
    Me.beveragestobuy_lbl.caption = "0"
    Me.frozenfoodstobuy_lbl.caption = "0"

    On Error GoTo 0
End Sub

Private Sub InitializeCategoryLabels()
    ' Sets default captions for category labels on the form
    On Error Resume Next

    Me.canMeatFishlbl.caption = "Meat & Fish"
    Me.canVegitableslbl.caption = "Canned Vegetables"
    Me.canFruitlbl.caption = "Canned Fruit"
    Me.DriedFoodlbl.caption = "Dried & Freeze Dried"
    Me.soupsLbl.caption = "Soups"
    Me.pastalbl.caption = "Pasta"
    Me.Legumeslbl.caption = "Legumes"
    Me.grainslbl.caption = "Grains"
    Me.oilsfatlbl.caption = "Oils & Fats"
    Me.dairyalternativeslbl.caption = "Dairy Alternatives"
    Me.bakinglbl.caption = "Baking"
    Me.snakstreatslbl.caption = "Snacks"
    Me.beverageslbl.caption = "Beverages"
    Me.frozentratslbl.caption = "Frozen"

    On Error GoTo 0
End Sub

Private Sub PopulateItemComboBox()
    ' Populates each combo box with items from specific ranges in the ItemList sheet
    ' Users can modify items directly on the ItemList sheet in the specified ranges
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim comboBoxes As Variant
    Dim ranges As Variant
    Dim i As Integer

    ' Reference the ItemList sheet
    Set ws = ThisWorkbook.Sheets("ItemList")

    ' Define combo boxes and their corresponding ranges from ItemList
    ' These ranges can be modified by users directly on the ItemList sheet
    comboBoxes = Array("cboMeatsFish", "cboCannedVeg", "cboCannedFruit", "cboDriedFrozen", _
                      "cboSoups", "cboPasta", "cboLegumes", "cboGrains", "cboOilFats", _
                      "cboDairyAlt", "cboBaking", "cboSnacks", "cboBeverages", "cboFrozen")

    ranges = Array("A2:A20", "A71:A91", "A142:A160", "A162:A178", _
                   "A180:A190", "A93:A103", "A22:A34", "A55:A69", "A105:A115", _
                   "A117:A129", "A36:A53", "A192:A207", "A131:A140", "A209:A224")

    ' Clear and populate each combo box
    For i = 0 To UBound(comboBoxes)
        PopulateComboBox Me.Controls(comboBoxes(i)), ws.Range(ranges(i))
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error populating item combo boxes: " & Err.description, vbCritical, "Population Error"
End Sub

Private Sub PopulateComboBox(cbo As ComboBox, dataRange As Range)
    ' Helper subroutine to clear and populate a single combo box with data from a range
    Dim cell As Range

    ' Clear existing items
    cbo.Clear
    
    ' Add empty option at the top
    cbo.AddItem ""

    ' Add items from range
    For Each cell In dataRange
        If Trim(cell.Value) <> "" Then
            cbo.AddItem cell.Value
        End If
    Next cell
End Sub

Private Function GetCurrentQuantity(itemName As String) As Double
    ' Retrieves the current quantity of an item from the StorageData sheet (Column C)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = ThisWorkbook.Sheets("StorageData")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = itemName Then
            GetCurrentQuantity = ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i

    GetCurrentQuantity = 0 ' Item not found
End Function

Private Function GetToBuyQuantity(itemName As String) As Double
    ' Retrieves the "To Buy" quantity of an item from the StorageData sheet (Column I)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = ThisWorkbook.Sheets("StorageData")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = itemName Then
            GetToBuyQuantity = ws.Cells(i, 9).Value ' Column I is the 9th column
            Exit Function
        End If
    Next i

    GetToBuyQuantity = 0 ' Item not found
End Function

' Combo box change events to update quantity labels and to buy labels
Private Sub cboMeatsFish_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboMeatsFish
    On Error Resume Next
    If Me.cboMeatsFish.Text <> "" Then
        currentMeatQty = GetCurrentQuantity(Me.cboMeatsFish.Text)
        Me.meatqtylbl.caption = Format(currentMeatQty, "0")
        Me.canmeatstobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboMeatsFish.Text), "0")
        UpdateCategoryLabel "canMeatFishlbl", Me.cboMeatsFish.Text
    Else
        Me.meatqtylbl.caption = "0"
        Me.canmeatstobuy_lbl.caption = "0"
        ResetCategoryLabel "canMeatFishlbl", "Meat & Fish"
    End If
    On Error GoTo 0
End Sub

Private Sub cboCannedVeg_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboCannedVeg
    On Error Resume Next
    If Me.cboCannedVeg.Text <> "" Then
        currentVegiQty = GetCurrentQuantity(Me.cboCannedVeg.Text)
        Me.vegiqtylbl.caption = Format(currentVegiQty, "0")
        Me.cannedvegtobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboCannedVeg.Text), "0")
        UpdateCategoryLabel "canVegitableslbl", Me.cboCannedVeg.Text
    Else
        Me.vegiqtylbl.caption = "0"
        Me.cannedvegtobuy_lbl.caption = "0"
        ResetCategoryLabel "canVegitableslbl", "Canned Vegetables"
    End If
    On Error GoTo 0
End Sub

Private Sub cboCannedFruit_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboCannedFruit
    On Error Resume Next
    If Me.cboCannedFruit.Text <> "" Then
        currentFruitQty = GetCurrentQuantity(Me.cboCannedFruit.Text)
        Me.fruitqtylbl.caption = Format(currentFruitQty, "0")
        Me.cannedfruittobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboCannedFruit.Text), "0")
        UpdateCategoryLabel "canFruitlbl", Me.cboCannedFruit.Text
    Else
        Me.fruitqtylbl.caption = "0"
        Me.cannedfruittobuy_lbl.caption = "0"
        ResetCategoryLabel "canFruitlbl", "Canned Fruit"
    End If
    On Error GoTo 0
End Sub

Private Sub cboDriedFrozen_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboDriedFrozen
    On Error Resume Next
    If Me.cboDriedFrozen.Text <> "" Then
        currentDriedQty = GetCurrentQuantity(Me.cboDriedFrozen.Text)
        Me.driedfoodqtylbl.caption = Format(currentDriedQty, "0")
        Me.driedfoodtobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboDriedFrozen.Text), "0")
        UpdateCategoryLabel "DriedFoodlbl", Me.cboDriedFrozen.Text
    Else
        Me.driedfoodqtylbl.caption = "0"
        Me.driedfoodtobuy_lbl.caption = "0"
        ResetCategoryLabel "DriedFoodlbl", "Dried & Freeze Dried"
    End If
    On Error GoTo 0
End Sub

Private Sub cboSoups_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboSoups
    On Error Resume Next
    If Me.cboSoups.Text <> "" Then
        currentSoupsQty = GetCurrentQuantity(Me.cboSoups.Text)
        Me.soupqtylbl.caption = Format(currentSoupsQty, "0")
        Me.soupstobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboSoups.Text), "0")
        UpdateCategoryLabel "soupsLbl", Me.cboSoups.Text
    Else
        Me.soupqtylbl.caption = "0"
        Me.soupstobuy_lbl.caption = "0"
        ResetCategoryLabel "soupsLbl", "Soups"
    End If
    On Error GoTo 0
End Sub

Private Sub cboPasta_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboPasta
    On Error Resume Next
    If Me.cboPasta.Text <> "" Then
        currentPastaQty = GetCurrentQuantity(Me.cboPasta.Text)
        Me.pastyqtylbl.caption = Format(currentPastaQty, "0")
        Me.pastatobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboPasta.Text), "0")
        UpdateCategoryLabel "pastalbl", Me.cboPasta.Text
    Else
        Me.pastyqtylbl.caption = "0"
        Me.pastatobuy_lbl.caption = "0"
        ResetCategoryLabel "pastalbl", "Pasta"
    End If
    On Error GoTo 0
End Sub

Private Sub cboLegumes_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboLegumes
    On Error Resume Next
    If Me.cboLegumes.Text <> "" Then
        currentLegumesQty = GetCurrentQuantity(Me.cboLegumes.Text)
        Me.legumesqtylbl.caption = Format(currentLegumesQty, "0")
        Me.legumestobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboLegumes.Text), "0")
        UpdateCategoryLabel "Legumeslbl", Me.cboLegumes.Text
    Else
        Me.legumesqtylbl.caption = "0"
        Me.legumestobuy_lbl.caption = "0"
        ResetCategoryLabel "Legumeslbl", "Legumes"
    End If
    On Error GoTo 0
End Sub

Private Sub cboGrains_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboGrains
    On Error Resume Next
    If Me.cboGrains.Text <> "" Then
        currentGrainsQty = GetCurrentQuantity(Me.cboGrains.Text)
        Me.grainsqtylbl.caption = Format(currentGrainsQty, "0")
        Me.grainstobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboGrains.Text), "0")
        UpdateCategoryLabel "grainslbl", Me.cboGrains.Text
    Else
        Me.grainsqtylbl.caption = "0"
        Me.grainstobuy_lbl.caption = "0"
        ResetCategoryLabel "grainslbl", "Grains"
    End If
    On Error GoTo 0
End Sub

Private Sub cboOilFats_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboOilFats
    On Error Resume Next
    If Me.cboOilFats.Text <> "" Then
        currentOilFatsQty = GetCurrentQuantity(Me.cboOilFats.Text)
        Me.oilqtylbl.caption = Format(currentOilFatsQty, "0")
        Me.oilstobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboOilFats.Text), "0")
        UpdateCategoryLabel "oilsfatlbl", Me.cboOilFats.Text
    Else
        Me.oilqtylbl.caption = "0"
        Me.oilstobuy_lbl.caption = "0"
        ResetCategoryLabel "oilsfatlbl", "Oils & Fats"
    End If
    On Error GoTo 0
End Sub

Private Sub cboDairyAlt_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboDairyAlt
    On Error Resume Next
    If Me.cboDairyAlt.Text <> "" Then
        currentDairyAltQty = GetCurrentQuantity(Me.cboDairyAlt.Text)
        Me.dairtyqtylbl.caption = Format(currentDairyAltQty, "0")
        Me.dairytobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboDairyAlt.Text), "0")
        UpdateCategoryLabel "dairyalternativeslbl", Me.cboDairyAlt.Text
    Else
        Me.dairtyqtylbl.caption = "0"
        Me.dairytobuy_lbl.caption = "0"
        ResetCategoryLabel "dairyalternativeslbl", "Dairy Alternatives"
    End If
    On Error GoTo 0
End Sub

Private Sub cboBaking_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboBaking
    On Error Resume Next
    If Me.cboBaking.Text <> "" Then
        currentBakingQty = GetCurrentQuantity(Me.cboBaking.Text)
        Me.bakingqtylbl.caption = Format(currentBakingQty, "0")
        Me.bakingtobut_lbl.caption = Format(GetToBuyQuantity(Me.cboBaking.Text), "0")
        UpdateCategoryLabel "bakinglbl", Me.cboBaking.Text
    Else
        Me.bakingqtylbl.caption = "0"
        Me.bakingtobut_lbl.caption = "0"
        ResetCategoryLabel "bakinglbl", "Baking"
    End If
    On Error GoTo 0
End Sub

Private Sub cboSnacks_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboSnacks
    On Error Resume Next
    If Me.cboSnacks.Text <> "" Then
        currentSnacksQty = GetCurrentQuantity(Me.cboSnacks.Text)
        Me.tratsqtylbl.caption = Format(currentSnacksQty, "0")
        Me.snacktobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboSnacks.Text), "0")
        UpdateCategoryLabel "snakstreatslbl", Me.cboSnacks.Text
    Else
        Me.tratsqtylbl.caption = "0"
        Me.snacktobuy_lbl.caption = "0"
        ResetCategoryLabel "snakstreatslbl", "Snacks"
    End If
    On Error GoTo 0
End Sub

Private Sub cboBeverages_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboBeverages
    On Error Resume Next
    If Me.cboBeverages.Text <> "" Then
        currentBevQty = GetCurrentQuantity(Me.cboBeverages.Text)
        Me.beveragesqtylbl.caption = Format(currentBevQty, "0")
        Me.beveragestobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboBeverages.Text), "0")
        UpdateCategoryLabel "beverageslbl", Me.cboBeverages.Text
    Else
        Me.beveragesqtylbl.caption = "0"
        Me.beveragestobuy_lbl.caption = "0"
        ResetCategoryLabel "beverageslbl", "Beverages"
    End If
    On Error GoTo 0
End Sub

Private Sub cboFrozen_Change()
    ' Updates the quantity label, category label, and to buy label when an item is selected in cboFrozen
    On Error Resume Next
    If Me.cboFrozen.Text <> "" Then
        currentFrozenQty = GetCurrentQuantity(Me.cboFrozen.Text)
        Me.Treatsqtylbl.caption = Format(currentFrozenQty, "0")
        Me.frozenfoodstobuy_lbl.caption = Format(GetToBuyQuantity(Me.cboFrozen.Text), "0")
        UpdateCategoryLabel "frozentratslbl", Me.cboFrozen.Text
    Else
        Me.Treatsqtylbl.caption = "0"
        Me.frozenfoodstobuy_lbl.caption = "0"
        ResetCategoryLabel "frozentratslbl", "Frozen"
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateCategoryLabel(labelName As String, selectedItem As String)
    ' Updates a category label with the item's category and changes color to indicate selection
    Dim categoryName As String
    On Error Resume Next

    categoryName = GetItemCategory(selectedItem)
    Me.Controls(labelName).caption = categoryName
    Me.Controls(labelName).ForeColor = RGB(0, 100, 0) ' Dark green to indicate selection
    On Error GoTo 0
End Sub

Private Sub ResetCategoryLabel(labelName As String, defaultCategory As String)
    ' Resets a category label to its default text and color
    On Error Resume Next
    Me.Controls(labelName).caption = defaultCategory
    Me.Controls(labelName).ForeColor = RGB(0, 0, 0) ' Black for default state
    On Error GoTo 0
End Sub

Private Sub cmdSave_Click()
    ' Saves all selected items from combo boxes to StorageData sheet by adding quantities from text boxes
    On Error GoTo ErrorHandler

    Dim comboBoxes As Variant
    Dim textBoxes As Variant
    Dim qtyLabels As Variant
    Dim toBuyLabels As Variant
    Dim selectedItem As String
    Dim selectedCategory As String
    Dim addQty As Double
    Dim confirmationMsg As String
    Dim i As Integer
    Dim savedCount As Integer

    ' Define combo boxes, quantity text boxes, quantity labels, and to buy labels
    comboBoxes = Array("cboMeatsFish", "cboCannedVeg", "cboCannedFruit", "cboDriedFrozen", _
                      "cboSoups", "cboPasta", "cboLegumes", "cboGrains", "cboOilFats", _
                      "cboDairyAlt", "cboBaking", "cboSnacks", "cboBeverages", "cboFrozen")

    textBoxes = Array("CanMeattext", "canvegtxt", "canfruittxt", "driedfoodtxt", _
                     "souptxt", "pastatxt", "legumestxt", "graintxt", "oiltxt", _
                     "dairytxt", "bakingtxt", "treatstxt", "beveragestxt", "frozentreatstxt")

    qtyLabels = Array("meatqtylbl", "vegiqtylbl", "fruitqtylbl", "driedfoodqtylbl", _
                      "soupqtylbl", "pastyqtylbl", "legumesqtylbl", "grainsqtylbl", "oilqtylbl", _
                      "dairtyqtylbl", "bakingqtylbl", "tratsqtylbl", "beveragesqtylbl", "Treatsqtylbl")

    toBuyLabels = Array("canmeatstobuy_lbl", "cannedvegtobuy_lbl", "cannedfruittobuy_lbl", "driedfoodtobuy_lbl", _
                       "soupstobuy_lbl", "pastatobuy_lbl", "legumestobuy_lbl", "grainstobuy_lbl", "oilstobuy_lbl", _
                       "dairytobuy_lbl", "bakingtobut_lbl", "snacktobuy_lbl", "beveragestobuy_lbl", "frozenfoodstobuy_lbl")

    confirmationMsg = ""
    savedCount = 0

    ' Loop through all combo boxes to save each selected item
    For i = 0 To UBound(comboBoxes)
        selectedItem = Me.Controls(comboBoxes(i)).Text
        If selectedItem <> "" Then
            ' Modified validation - allow blank quantities (default to 0)
            If Me.Controls(textBoxes(i)).Text = "" Then
                addQty = 0
            Else
                addQty = Val(Me.Controls(textBoxes(i)).Text)
                If addQty < 0 Then
                    MsgBox "Please enter a valid quantity (0 or greater) for " & selectedItem & ".", vbExclamation, "Validation Error"
                    Me.Controls(textBoxes(i)).SetFocus
                    Exit Sub
                End If
            End If

            ' Get the category for the selected item
            selectedCategory = GetItemCategory(selectedItem)

            ' Update inventory (only if quantity > 0)
            If addQty > 0 Then
                Call UpdateInventory(selectedItem, addQty, selectedCategory)
            End If

            ' Update the corresponding quantity label and to buy label
            Me.Controls(qtyLabels(i)).caption = Format(GetCurrentQuantity(selectedItem), "0")
            Me.Controls(toBuyLabels(i)).caption = Format(GetToBuyQuantity(selectedItem), "0")

            ' Build confirmation message
            If addQty > 0 Then
                confirmationMsg = confirmationMsg & "Item '" & selectedItem & "' saved successfully!" & vbCrLf & _
                                  "Quantity Added: " & addQty & vbCrLf & _
                                  "New Total Quantity: " & GetCurrentQuantity(selectedItem) & vbCrLf & _
                                  "Category: " & selectedCategory & vbCrLf & vbCrLf
            Else
                confirmationMsg = confirmationMsg & "Item '" & selectedItem & "' selected (no quantity added)." & vbCrLf & _
                                  "Current Total Quantity: " & GetCurrentQuantity(selectedItem) & vbCrLf & _
                                  "Category: " & selectedCategory & vbCrLf & vbCrLf
            End If
            savedCount = savedCount + 1
        End If
    Next i

    ' Show confirmation if any items were processed
    If savedCount > 0 Then
        MsgBox confirmationMsg, vbInformation, "Save Complete (" & savedCount & " items processed)"
    Else
        MsgBox "No items selected to save.", vbInformation, "Save Complete"
    End If
    
    ' Unhide the StorageData worksheet
    Call ShowStorageData
    ' Clear form for next entry
    Call ClearForm

    Exit Sub

ErrorHandler:
    MsgBox "Error saving items: " & Err.description, vbCritical, "Save Error"
End Sub

Private Sub UpdateInventory(itemName As String, addQuantity As Double, categoryName As String)
    ' Updates or adds an item's quantity in the StorageData sheet
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim found As Boolean

    Set ws = ThisWorkbook.Sheets("StorageData")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    found = False

    ' Look for existing item
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = itemName Then
            ' Update existing item
            ws.Cells(i, 3).Value = ws.Cells(i, 3).Value + addQuantity
            ws.Cells(i, 4).Value = Date ' Update purchase date
            found = True
            Exit For
        End If
    Next i

    ' If item not found, add new row
    If Not found Then
        lastRow = lastRow + 1
        ws.Cells(lastRow, 1).Value = itemName
        ws.Cells(lastRow, 2).Value = categoryName
        ws.Cells(lastRow, 3).Value = addQuantity
        ws.Cells(lastRow, 4).Value = Date
    End If
End Sub

Private Sub cmdRefresh_Click()
    ' Refreshes the combo boxes with latest items from ItemList and resets labels
    PopulateItemComboBox
    InitializeCategoryLabels
    InitializeToBuyLabels
    MsgBox "Item lists refreshed!", vbInformation, "Refresh Complete"
End Sub

Private Sub cmdAddNewItem_Click()
    ' Adds a new item to the ItemList sheet and refreshes combo boxes
    On Error GoTo ErrorHandler

    Dim newItem As String
    Dim selectedCategory As String
    Dim ws As Worksheet
    Dim nextRow As Long

    newItem = InputBox("Enter new item name:", "Add New Item")

    If newItem <> "" Then
        selectedCategory = InputBox("Enter category for this item:", "Item Category")

        If selectedCategory <> "" Then
            Set ws = ThisWorkbook.Sheets("ItemList")
            nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

            ws.Cells(nextRow, 1).Value = newItem
            ws.Cells(nextRow, 2).Value = selectedCategory

            ' Refresh the combo boxes
            PopulateItemComboBox

            MsgBox "New item '" & newItem & "' added successfully!", vbInformation, "Item Added"
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error adding new item: " & Err.description, vbCritical, "Add Item Error"
End Sub

Private Function GetItemCategory(itemName As String) As String
    ' Returns the category for a given item based on predefined mappings
    Select Case LCase(Trim(itemName))
        Case "anchovies", "beef (chunks, roast beef)", "chicken", "clams", "corned beef & corned beef hash", _
             "crab meat", "ham (diced, deviled)", "liver pâté", "mackerel", "oysters", _
             "pork (chunks, pulled pork)", "potted meat spread", "salmon", "sardines", "scallops", _
             "shrimp", "spam/luncheon meat", "tuna", "turkey", "vienna sausages"
            GetItemCategory = "Meat & Fish"
        Case "artichoke hearts", "asparagus", "bamboo shoots", "beets", "carrots", _
             "corn (whole kernel, creamed)", "green beans (cut, french-style)", "hearts of palm", _
             "hominy", "leafy greens (turnip, collard, mustard)", "mushrooms", "okra", "peas", _
             "potatoes (whole, sliced, diced)", "pumpkin puree", "sauerkraut", "spinach", _
             "tomato sauce/paste", "tomatoes (whole, diced, crushed, stewed)", "water chestnuts", _
             "yams / sweet potatoes"
            GetItemCategory = "Canned Vegetables"
        Case "applesauce", "apricots", "cherries", "cranberry sauce", _
             "fruit cocktail", "grapefruit", "mandarin oranges", "mango", "peaches", "pears", _
             "pineapple", "plums"
            GetItemCategory = "Canned Fruit"
        Case "freeze-dried meats (beef, chicken, turkey)", "dehydrated potatoes (flakes, slices)", _
             "freeze-dried fruits", "freeze-dried vegetables", "dried fruit"
            GetItemCategory = "Dried & Freeze Dried"
        Case "canned soups (vegetable, chicken noodle, tomato, etc.)", "canned curry meals"
            GetItemCategory = "Soups"
        Case "egg noodles", "lasagna noodles", "macaroni", "ramen noodles", "rice noodles", "spaghetti"
            GetItemCategory = "Pasta"
        Case "black beans", "garbanzo beans (chickpeas)", "kidney beans", "lentils (green, red, brown)", _
             "navy beans", "pinto beans", "split peas", "dehydrated beans"
            GetItemCategory = "Legumes"
        Case "all-purpose flour", "bread flour", "cornmeal", "whole wheat flour", "barley", _
             "brown rice (shorter shelf life)", "bulgur wheat", "farro", "millet", "quick oats", _
             "quinoa", "rolled oats", "steel-cut oats", "white rice (long-term storage friendly)"
            GetItemCategory = "Grains"
        Case "coconut oil", "ghee (clarified butter)", "olive oil", "peanut oil", "shortening", "vegetable oil"
            GetItemCategory = "Oils & Fats"
        Case "canned evaporated milk", "canned sweetened condensed milk", "powdered coffee creamer", _
             "powdered milk", "shelf-stable plant milks (soy, almond, oat)", "uht milk (shelf-stable cartons)"
            GetItemCategory = "Dairy Alternatives"
        Case "baking powder", "baking soda", "cocoa powder", "salt (iodized & non-iodized)", "yeast", _
             "honey", "maple syrup", "molasses", "sugar (white, brown, powdered)"
            GetItemCategory = "Baking"
        Case "crackers", "granola bars", "popcorn kernels", "trail mix", "hard candy", _
             "chocolate (dark lasts longest)", "peanut butter"
            GetItemCategory = "Snacks"
        Case "coffee (instant, whole bean, ground)", "hot cocoa mix", _
             "powdered drink mixes (lemonade, sports drinks)", "tea (loose leaf, bags)"
            GetItemCategory = "Beverages"
        Case "hamburger", "hotdogs", "stew meat", "meatballs", "chicken breasts", "chicken thighs", _
             "ham", "corndogs", "chicken nuggets", "ice cream"
            GetItemCategory = "Frozen"
        Case Else
            GetItemCategory = "Unknown Category"
    End Select
End Function

Private Sub ClearForm()
    ' Clears all combo boxes, quantity text boxes, resets category labels, and resets to buy labels for the next entry
    Dim comboBoxes As Variant
    Dim textBoxes As Variant
    Dim i As Integer
    
    ' Define combo boxes and their corresponding quantity text boxes
    comboBoxes = Array("cboMeatsFish", "cboCannedVeg", "cboCannedFruit", "cboDriedFrozen", _
                      "cboSoups", "cboPasta", "cboLegumes", "cboGrains", "cboOilFats", _
                      "cboDairyAlt", "cboBaking", "cboSnacks", "cboBeverages", "cboFrozen")
    
    textBoxes = Array("CanMeattext", "canvegtxt", "canfruittxt", "driedfoodtxt", _
                     "souptxt", "pastatxt", "legumestxt", "graintxt", "oiltxt", _
                     "dairytxt", "bakingtxt", "treatstxt", "beveragestxt", "frozentreatstxt")
    
    ' Clear all combo boxes - set to empty option
    For i = 0 To UBound(comboBoxes)
        On Error Resume Next
        Me.Controls(comboBoxes(i)).Text = ""
        If Err.Number <> 0 Then
            Debug.Print "Error clearing combo box " & comboBoxes(i) & ": " & Err.description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    ' Clear all quantity text boxes
    For i = 0 To UBound(textBoxes)
        On Error Resume Next
        Me.Controls(textBoxes(i)).Text = ""
        If Err.Number <> 0 Then
            Debug.Print "Error clearing text box " & textBoxes(i) & ": " & Err.description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    ' Reset all category labels and to buy labels to their default state
    Call InitializeCategoryLabels
    Call InitializeToBuyLabels
    
    ' Reset Me.Tag
    Me.Tag = ""
End Sub
' Show the shopping list worksheet
Private Sub Submitclk_Click()
    Call ImprovedShoppingListMacro_Simplified
    Unload Me
    
    Call ShowShoppingList
    
End Sub


