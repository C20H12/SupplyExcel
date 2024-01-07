Sub master()
    ThisWorkbook.Sheets("Master").Range("A3").Value = "Generating"
    Dim lastRow As Long
    lastRow = ActiveSheet.UsedRange.Rows.count
    Dim i As Integer
    For i = 3 To lastRow
        If Not i = lastRow Then
            ThisWorkbook.Sheets("Master").Rows(4).Delete
        Else
             ThisWorkbook.Sheets("Master").Rows(3).Delete
        End If
    Next i
    
    Dim OrigSheet As Worksheet
    Set OrigSheet = ActiveSheet
    
    Dim ws As Worksheet
    
    Dim PickUpSheetRow As Integer
    PickUpSheetRow = 2
    
    ActiveWorkbook.Worksheets("Master").ListObjects("StatusTable").AutoFilter.ShowAllData
    
   ' remove all buttons so that there is no overlap
    For Each btn In OrigSheet.Buttons
        If btn.Caption <> "Generate" Then
            btn.Delete
        End If
    Next btn

    For Each ws In ThisWorkbook.Worksheets
        ' ignore special sheets
        If isSpecialSheet(ws.Name) Then
            GoTo continue
        End If
        
        ' get all the items' nsn, size, and status in lists
        Dim nsnRange As Range
        Dim cell As Range
        Dim sizes() As Variant
        Dim status() As Variant
        Dim count As Integer
        count = 0
        Set nsnRange = ws.Range("A6:A26")
        For Each cell In nsnRange
            ReDim Preserve sizes(count)
            sizes(count) = cell.Offset(0, 4).Value
            ReDim Preserve status(count)
            status(count) = cell.Offset(0, 6).Value
            count = count + 1
        Next cell
        
        ' get name
        OrigSheet.Cells(PickUpSheetRow + 1, 1).Value = ws.Range("C2").Value + ", " + ws.Range("E2").Value
        Dim linkAddress As String
        linkAddress = "" & ws.Name & "!A1"
        OrigSheet.Hyperlinks.Add Anchor:=OrigSheet.Cells(PickUpSheetRow + 1, 1), _
                          Address:="", _
                          SubAddress:=linkAddress, _
                          TextToDisplay:=ws.Range("C2").Value + ", " + ws.Range("E2").Value
                      
    
        Dim hasIncomplete As Boolean
        hasIncomplete = False

        For i = 0 To 20
        
            If i = 9 Or i = 14 Or Len(Trim(sizes(i))) = 0 Then
                GoTo continueinner
            End If
            
            ' fill in the size
            OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Value = sizes(i)
            
            ' highlight
            If status(i) = "UNP" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(255, 117, 117)
            ElseIf status(i) = "In Stock" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(251, 163, 251)
            ElseIf status(i) = "Pick Up" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(146, 208, 80)
            ElseIf status(i) = "Ready To Order" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(246, 246, 106)
            ElseIf status(i) = "Ordered" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(244, 176, 132)
            ElseIf status(i) = "Complete" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(155, 194, 230)
            ElseIf status(i) = "Returned" Then
              OrigSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(128, 128, 128)
            End If
            
            If status(i) <> "Complete" Then
                hasIncomplete = True
            End If
            
continueinner:
        Next i
        
        
        ' highlight name as red if not all complete
        If hasIncomplete Then
            OrigSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = RGB(252, 136, 136)
        Else
            OrigSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = RGB(140, 255, 140)
        End If
        
        Dim t As Range
        Set t = OrigSheet.Cells(PickUpSheetRow + 1, 23)
        Set btn = ActiveSheet.Buttons.Add(t.left, t.Top, t.Width, t.Height)
        With btn
          .OnAction = "'togglePersonAsComplete " & PickUpSheetRow + 1 & "'"
          .Caption = "Toggle"
          .Name = "Toggle" & PickUpSheetRow 'need to have a unique name or else it won't delete
        End With
        
        PickUpSheetRow = PickUpSheetRow + 1
        
continue:
    Next ws
    
End Sub

Sub togglePersonAsComplete(r As Integer)
    If ActiveSheet.Cells(r, 1).Interior.Color = RGB(140, 255, 140) Then
        ActiveSheet.Cells(r, 1).Interior.Color = RGB(253, 234, 93)
    ElseIf ActiveSheet.Cells(r, 1).Interior.Color = RGB(253, 234, 93) Then
         ActiveSheet.Cells(r, 1).Interior.Color = RGB(252, 136, 136)
    Else:
        ActiveSheet.Cells(r, 1).Interior.Color = RGB(140, 255, 140)
    End If
End Sub