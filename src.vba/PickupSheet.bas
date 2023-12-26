Sub pickup()

    Dim origSheet As Worksheet
    Set origSheet = ActiveSheet
    
    ' open the inventory book for checking stock
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Supply_Physical_Inventory.xlsx")
    
    Dim ws As Worksheet
    
    Dim PickUpSheetRow As Integer
    PickUpSheetRow = 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Menu" Or ws.Name = "Importing" Or ws.Name = "Pickup" Or ws.Name = "Template" Then
            GoTo continue
        End If
        
        Dim nsnRange As Range
        Dim cell As Range
        Dim stockNumbers() As Variant
        Dim sizes() As Variant
        Dim count As Integer
        count = 0
        Set nsnRange = ws.Range("A6:A24")
        For Each cell In nsnRange
            ReDim Preserve stockNumbers(count)
            stockNumbers(count) = cell.Value
            ReDim Preserve sizes(count)
            sizes(count) = cell.Offset(0, 4).Value
            count = count + 1
        Next cell
        
        origSheet.Cells(PickUpSheetRow + 1, 1).Value = ws.Range("C2").Value + ", " + ws.Range("E2").Value

        For i = 0 To 16
            If i = 9 Or i = 14 Then
                GoTo continueInner
            End If
            
            If Len(Trim(stockNumbers(i))) = 0 Then
                origSheet.Cells(PickUpSheetRow + 1, i + 2).Value = "NO SIZE"
                GoTo continueInner
            End If
            
            ' find the right nsn inside the inventory sheet and store it here
            Dim Loc As Range
            
            For Each sh In wb.Worksheets
                With sh.UsedRange
                    Set Loc = .Cells.Find(What:=stockNumbers(i))
                    If Not Loc Is Nothing Then
                        Exit For
                    End If
                End With
                Set Loc = Nothing
            Next
            
            If Loc Is Nothing Then
                Exit Sub
            End If
            
        
            Dim Row As Integer
            Dim Col As Integer
        
            Row = Loc.Row
            Col = Loc.Column
            
            Dim QTYcol As Integer
            
            ' find the quantity
            For j = Col To Col + 8
                If Loc.Worksheet.Cells(3, j).Value = "QTY" Then
                    QTYcol = j
                    Exit For
                End If
            Next j
            
            origSheet.Cells(PickUpSheetRow + 1, i + 2).Value = sizes(i)
            
            If Loc.Worksheet.Cells(Row, QTYcol).Value <> 0 Then
                origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(176, 255, 177)
            Else
                origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(255, 176, 177)
            End If
            
continueInner:
        Next i
        
        PickUpSheetRow = PickUpSheetRow + 1
        
continue:
    Next ws
    
    wb.Close
End Sub