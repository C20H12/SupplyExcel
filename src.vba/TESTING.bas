Sub Testing()
    

    
    
    ' Dim wb As Workbook

    ' Set wb = Workbooks.Open(ThisWorkbook.Path & "\Supply_Physical_Inventory.xlsx")

    ' Dim Loc As Range
    
    ' For Each Sh In wb.Worksheets
    
    '     With Sh.UsedRange
    '         Set Loc = .Cells.Find(What:="8410-21-912-3645")
    '         If Not Loc Is Nothing Then
            
                
                
    '             Exit For
    '         End If
    '     End With
    '     Set Loc = Nothing
    ' Next
    

    ' Dim Row As Integer
    ' Dim Col As Integer

    ' Row = Loc.Row
    ' Col = Loc.Column
    
    ' Debug.Print Loc.Worksheet.Name, Row, Col
    
    ' For i = Col To Col + 8
    '     Debug.Print Loc.Worksheet.Cells(3, i).Value
    '     If Loc.Worksheet.Cells(3, i).Value = "QTY" Then
    '         Debug.Print Loc.Worksheet.Cells(Row, i).Value
    '         Loc.Worksheet.Cells(Row, i).Value = 111
    '         Exit For
    '     End If
    ' Next i
    
    ' Debug.Print ThisWorkbook.Sheets("Importing").UsedRange.Rows.Count
    
    ' Debug.Print "8410-21-912-3651" Like "####*-##-###-####"
    ' Debug.Print "18410-21-912-3651" Like "####*-##-###-####"
    ' Debug.Print "a8410a-21-912-3651" Like "####*-##-###-####"


'    Debug.Print GetNSNFromSize("Tunic", "8062", True)
'    Debug.Print GetNSNFromSize("Tunic", "9999", True)
'    Debug.Print GetNSNFromSize("Collar Shirt", "13", True)
'    Debug.Print GetNSNFromSize("Parka", "6432", True)

    Dim exchangeHistoryRow() As Variant
    With Sheets("Wendy_Lee_5804A09C")
        exchangeHistoryRow = Array(Range("A38"), Range("B38"), Range("C38"), Range("D38"))
        Debug.Print .ListObjects(.name & "ExchangeTable").Range.Rows.count
    End With

    Debug.Print Join(exchangeHistoryRow, "===")
    
    
    
End Sub
