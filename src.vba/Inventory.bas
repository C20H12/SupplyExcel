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
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Supply_Physical_Inventory.xlsx")

    ' find the right nsn inside the inventory sheet and store it here
    Dim Loc As Range
    
    For Each Sh In wb.Worksheets
        With Sh.UsedRange
            Set Loc = .Cells.Find(What:=nsn)
            If Not Loc Is Nothing Then
                Exit For
            End If
        End With
        Set Loc = Nothing
    Next
    
    If Loc Is Nothing Then
        MsgBox "Selected value not found."
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
    
    If Not Modified Then
        Loc.Worksheet.Cells(Row, QTYcol).Value = Modified
    End If
End Sub