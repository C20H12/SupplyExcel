Sub pickup()
  ThisWorkbook.Sheets("Pickup").Range("A2").Value = "Generating"
    Dim lastRow As Long
    lastRow = ActiveSheet.UsedRange.Rows.count
    Dim i As Integer
    For i = 3 To lastRow
        If Not i = lastRow Then
            ThisWorkbook.Sheets("Pickup").Rows(3).Delete
        Else
             ThisWorkbook.Sheets("Pickup").Rows(2).Delete
        End If
    Next i
    
    Dim OrigSheet As Worksheet
    Set OrigSheet = ActiveSheet
    
    Dim ws As Worksheet
    
    Dim PickUpSheetRow As Integer
    PickUpSheetRow = 1
    
    ' remove all buttons so that there is no overlap
    For Each btn In OrigSheet.Buttons
        If btn.Caption <> "Generate" Then
            btn.Delete
        End If
    Next btn

    For Each ws In ThisWorkbook.Worksheets
        ' ignore special sheets
        If isSpecialSheet(ws.Name) Then
            GoTo Continue
        End If
        
        ' get all the items' nsn, size, and status in lists
        Dim nsnRange As Range
        Dim cell As Range
        Dim sizes() As Variant
        Dim status() As Variant
        Dim hasReadyToPickUp As Boolean
        hasReadyToPickUp = False
        Dim count As Integer
        count = 0
        Set nsnRange = ws.Range("A6:A26")
        For Each cell In nsnRange
            ReDim Preserve sizes(count)
            sizes(count) = cell.Offset(0, 4).Value
            ReDim Preserve status(count)
            status(count) = cell.Offset(0, 6).Value
            If status(count) = "Pick Up" Then
                hasReadyToPickUp = True
            End If
            count = count + 1
        Next cell
        
        If Not hasReadyToPickUp Then
            GoTo Continue
        End If
        
        ' get name
        OrigSheet.Cells(PickUpSheetRow + 1, 1).Value = ws.Range("C2").Value + ", " + ws.Range("E2").Value
        Dim linkAddress As String
        linkAddress = "" & ws.Name & "!A1"
        OrigSheet.Hyperlinks.Add Anchor:=OrigSheet.Cells(PickUpSheetRow + 1, 1), _
                          Address:="", _
                          SubAddress:=linkAddress, _
                          TextToDisplay:=ws.Range("C2").Value + ", " + ws.Range("E2").Value
        For i = 0 To 20
            If i = 9 Or i = 14 Or Len(Trim(sizes(i))) = 0 Or status(i) <> "Pick Up" Then
                GoTo continueinner
            End If
            
            ' fill in the size
            OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Value = sizes(i)
            
            ' highlight if ready to pick up
            OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(176, 255, 177)
            
continueinner:
        Next i
        
        Dim t As Range
        Set t = OrigSheet.Cells(PickUpSheetRow + 1, 24)
        Set btn = ActiveSheet.Buttons.Add(t.left, t.Top, t.Width, t.Height)
        Dim SheetName As String
        SheetName = """" & ws.Name & """"
        With btn
          .OnAction = "'markPickUpAsComplete " & SheetName & ", " & PickUpSheetRow & "'"
          .Caption = "Complete"
          .Name = "Complete" & PickUpSheetRow
        End With
        
        PickUpSheetRow = PickUpSheetRow + 1
        
Continue:
    Next ws
    
End Sub

Sub markPickUpAsComplete(n As String, r As Integer)
    If MsgBox("Are you sure you want to perform this action?", vbYesNo) = vbNo Then
            Exit Sub
    End If
    
    For Each cell In ActiveWorkbook.Worksheets(n).Range("G6:G24")
        If cell.Value = "Pick Up" Then
            cell.Value = "Complete"
        End If
    Next cell
    Rows(r + 1).EntireRow.Delete
End Sub