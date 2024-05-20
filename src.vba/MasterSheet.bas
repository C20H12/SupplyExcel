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
    
    Dim origSheet As Worksheet
    Set origSheet = ActiveSheet
    
    Dim ws As Worksheet
    
    Dim PickUpSheetRow As Integer
    PickUpSheetRow = 2
    
    ' remove all buttons so that there is no overlap
    For Each btn In origSheet.Buttons
        If btn.Caption <> "Generate" Then
            btn.Delete
        End If
    Next btn
    
    ActiveWorkbook.Worksheets("Master").ListObjects("StatusTable").AutoFilter.ShowAllData
    
    Dim sheetCount As Integer
    sheetCount = 0

    For Each ws In ThisWorkbook.Worksheets
        
        ' ignore special sheets
        If isSpecialSheet(ws.name) Then
            GoTo continue
        End If
        
        sheetCount = sheetCount + 1
        
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
        linkAddress = "" & ws.name & "!A1"
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
            Else
              origSheet.Cells(PickUpSheetRow + 1, i + 2).Interior.Color = RGB(231, 230, 230)
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
        If Not IsStringEmpty(ws.Cells(27, 1).Value) Then
            origSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = ws.Cells(27, 1).Value
        Else
        
            If hasIncomplete Then
                origSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = RGB(252, 136, 136)
                ws.Cells(27, 1).Value = RGB(252, 136, 136)
            Else
                origSheet.Cells(PickUpSheetRow + 1, 1).Interior.Color = RGB(140, 255, 140)
                ws.Cells(27, 1).Value = RGB(140, 255, 140)
            End If
        
        End If
        
        Dim t As Range
        Set t = origSheet.Cells(PickUpSheetRow + 1, 23)
        Set btn = ActiveSheet.Buttons.Add(t.left + 1, t.Top + 1, t.Width - 2, t.Height - 2)
        
        Dim actionstring As String
        actionstring = "togglePersonAsComplete " & PickUpSheetRow + 1 & ", " & """" & ws.name & """"
        
        With btn
            .OnAction = "'" & actionstring & "'"
            .Text = "Toggle" & PickUpSheetRow 'need to have a unique name or else it won't delete
        End With
        
        
        PickUpSheetRow = PickUpSheetRow + 1
        
continue:
    Next ws
    
End Sub

Sub togglePersonAsComplete(r As Integer, name As String)
    If ActiveSheet.Cells(r, 1).Interior.Color = RGB(140, 255, 140) Then
        ActiveSheet.Cells(r, 1).Interior.Color = RGB(253, 234, 93)
        Sheets(name).Cells(27, 1).Value = RGB(253, 234, 93)
    ElseIf ActiveSheet.Cells(r, 1).Interior.Color = RGB(253, 234, 93) Then
         ActiveSheet.Cells(r, 1).Interior.Color = RGB(252, 136, 136)
         Sheets(name).Cells(27, 1).Value = RGB(252, 136, 136)
    Else:
        ActiveSheet.Cells(r, 1).Interior.Color = RGB(140, 255, 140)
        Sheets(name).Cells(27, 1).Value = RGB(140, 255, 140)
    End If
End Sub
