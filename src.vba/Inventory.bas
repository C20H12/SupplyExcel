Sub SelectFileAndStorePath()
    Dim selectedFile As Variant
    
    ' Open the file explorer and prompt the user to select a file
    selectedFile = Application.GetOpenFilename("All Files (*.*), *.*", Title:="Select a File")
    
    ' Check if the user selected a file
    If selectedFile <> "False" Then
        ' Store the file path in cell A1 of the "Menu" sheet
        Worksheets("Menu").Range("CW1").Value = selectedFile
    Else
        ' User canceled the file selection
        MsgBox "File selection canceled."
    End If
End Sub


Function FindInInventory(nsn As String, Optional closeAfter As Boolean = False) As Integer
    Dim wb As Workbook
    ' get object makes it not show up
    On Error GoTo Copy
    Set wb = GetObject(ThisWorkbook.Path & "\Supply_Physical_Inventory.xlsm")
Copy:
     Set wb = Workbooks.Open(ThisWorkbook.Path & "\Supply_Physical_Inventory.xlsm")
    
    ' find the right nsn inside the inventory sheet and store it here
    
    Dim Loc As Range
    
    For Each sh In wb.Worksheets
        With sh.UsedRange
            Set Loc = .Cells.Find(What:=nsn)
            If Not Loc Is Nothing Then
                Exit For
            End If
        End With
        Set Loc = Nothing
    Next
    
    If Loc Is Nothing Then
        FindInInventory = -999
        If closeAfter Then
            wb.Close savechanges:=False
        End If
        Exit Function
    End If
    

    Dim Row As Integer
    Dim Col As Integer

    Row = Loc.Row
    Col = Loc.Column
    
    Dim QTYcol As Integer
    
    For i = Col To Col + 8
        'Debug.Print Loc.Worksheet.Cells(3, i).Value
        If Loc.Worksheet.Cells(3, i).Value = "QTY" Then
            'Debug.Print Loc.Worksheet.Cells(Row, i).Value
            QTYcol = i
            Exit For
        End If
    Next i
    
    FindInInventory = CInt(Loc.Worksheet.Cells(Row, QTYcol).Value)
    
    If closeAfter Then
        wb.Close savechanges:=False
    End If
    

End Function

Sub UpdateInStockStatus()
    For i = 6 To 24
        Dim nsn As String
        nsn = ActiveSheet.Range("A" & i).Value
        Dim status As String
        status = ActiveSheet.Range("G" & i).Value
                
        If Not IsStringEmpty(nsn) And status = "UNP" Then
            If FindInInventory(nsn) > 0 Then
                ActiveSheet.Range("G" & i).Value = "In Stock"
            End If
        End If
    Next i
End Sub

Sub InventoryInteract()
    ' get the selected nsn
    Dim nsn As String
    nsn = ActiveCell.Value
    
    If Not nsn Like "####*-##-###-####" Then
        MsgBox "Selected value is not a NSN"
        Exit Sub
    End If
    
    
    ' open the inventory book for modifying
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Supply_Physical_Inventory.xlsm")

    ' find the right nsn inside the inventory sheet and store it here
    Dim Loc As Range
    
    For Each sh In wb.Worksheets
        With sh.UsedRange
            Set Loc = .Cells.Find(What:=nsn)
            If Not Loc Is Nothing Then
                Exit For
            End If
        End With
        Set Loc = Nothing
    Next
    
    If Loc Is Nothing Then
        MsgBox "Selected value not found."
        wb.Close
        Exit Sub
    End If
    

    Dim Row As Integer
    Dim Col As Integer

    Row = Loc.Row
    Col = Loc.Column
    
    Dim QTYcol As Integer
    
    For i = Col To Col + 8
        'Debug.Print Loc.Worksheet.Cells(3, i).Value
        If Loc.Worksheet.Cells(3, i).Value = "QTY" Then
            'Debug.Print Loc.Worksheet.Cells(Row, i).Value
            QTYcol = i
            Exit For
        End If
    Next i
    
    Dim Modified As Variant
    Modified = Application.InputBox("Modify the quantity of this item:", "Inventory", Loc.Worksheet.Cells(Row, QTYcol).Value, Type:=1)
    
    If Not Modified = False Then
        Loc.Worksheet.Cells(Row, QTYcol).Value = Modified
    End If
    
    wb.Close savechanges:=True
End Sub