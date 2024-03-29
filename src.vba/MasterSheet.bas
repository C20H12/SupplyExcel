Sub master()
    ThisWorkbook.Sheets("Master").Range("A3").Value = "Generating"
    Dim lastRow As Long
    lastRow = ActiveSheet.UsedRange.Rows.count
    Dim nameStatus() As Variant
    Dim i As Integer
    For i = 3 To lastRow
        'store the color on the name cell (status set by the user)
        ReDim Preserve nameStatus(i - 2)
        If Not i = lastRow Then
            nameStatus(i - 2) = ThisWorkbook.Sheets("Master").Cells(4, 1).Interior.Color
            ThisWorkbook.Sheets("Master").Rows(4).Delete
        Else
            nameStatus(0) = ThisWorkbook.Sheets("Master").Cells(3, 1).Interior.Color
            ThisWorkbook.Sheets("Master").Rows(3).Delete
        End If
    Next i
    
    Dim origSheet As Worksheet
    Set origSheet = ActiveSheet
    
    Dim ws As Worksheet
    
    Dim PickUpSheetRow As Integer
    PickUpSheetRow = 2
    
    ActiveWorkbook.Worksheets("Master").ListObjects("StatusTable").AutoFilter.ShowAllData
    
   ' remove all buttons so that there is no overlap
    For Each btn In origSheet.Buttons
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
        origSheet.Cells(PickUpSheetRow + 1, 1).Value = ws.Range("C2").Value + ", " + ws.Range("E2").Value
        Dim linkAddress As String
        linkAddress = "" & ws.Name & "!A1"
        origSheet.Hyperlinks.Add Anchor:=origSheet.Cells(PickUpSheetRow + 1, 1), _
                          Address:="", _
                          SubAddress:=linkAddress, _
                          TextToDisplay:=ws.Range("C2").Value + ", " + ws.Range("E2").Value
                      
    
        Dim hasIncomplete As Boolean
        hasIncomplete = False
        Dim statusString As String
        statusString = ""

        For i = 0 To 20
        
            If i = 9 Or i = 14 Then
                GoTo continueinner
            End If
            
            ' fill in the size
            origSheet.Cells(PickUpSheetRow + 1, i + 2).Value = sizes(i)
            
            ' highlight
            If status(i) = "UNP" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(255, 117, 117)
            ElseIf status(i) = "In Stock" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(251, 163, 251)
            ElseIf status(i) = "Pick Up" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(146, 208, 80)
            ElseIf status(i) = "Ready To Order" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(246, 246, 106)
            ElseIf status(i) = "Ordered" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(244, 176, 132)
            ElseIf status(i) = "Complete" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(155, 194, 230)
            ElseIf status(i) = "Returned" Then
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(128, 128, 128)
            End If
            
            If status(i) <> "Complete" Then
                hasIncomplete = True
            End If

            If Not InStr(statusString, status(i)) > 0 Then
                statusString = status(i) & "," & statusString
            End If
            
continueinner:
        Next i
        
        origSheet.Cells(PickUpSheetRow + 1, 24).Value = statusString
        
        ' highlight name as red if not all complete
        If hasIncomplete Then
            origSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = RGB(252, 136, 136)
        Else
            origSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = RGB(140, 255, 140)
        End If
        
        Dim t As Range
        Set t = origSheet.Cells(PickUpSheetRow + 1, 23)
        Set btn = ActiveSheet.Buttons.Add(t.left, t.Top, t.Width, t.Height)
        With btn
          .OnAction = "'togglePersonAsComplete " & PickUpSheetRow + 1 & "'"
          .Caption = "Toggle"
          .Name = "Toggle" & PickUpSheetRow 'need to have a unique name or else it won't delete
        End With
        
        PickUpSheetRow = PickUpSheetRow + 1
        
continue:
    Next ws
    
    ' restore color
    For i = 0 To lastRow - 3
        origSheet.Cells(i + 3, 1).Interior.Color = nameStatus(i)
    Next i
    
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