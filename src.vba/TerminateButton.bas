Sub MarkAsSOS()

    Dim ws As Worksheet
    Dim currentID As Variant
    Dim menuSheet As Worksheet
    Dim menuTable As ListObject
    Dim idColumn As ListColumn
    Dim deleteCell As Range
    
    ' Get reference to the current sheet
    Set ws = ActiveSheet
    
    If ws.Name = "Template" Then
        MsgBox "Cannot mark the template as SOS"
        Exit Sub
    Else
        If MsgBox("Are you sure you want to perform this action?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ws.Range("G27:G28").Merge
    ws.Range("G27").Value = "S.O.S"
    
    With ws.Range("G27")
        .Font.Size = 25
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
End Sub


Sub Terminate()

    Dim ws As Worksheet
    Dim currentID As Variant
    Dim menuSheet As Worksheet
    Dim menuTable As ListObject
    Dim idColumn As ListColumn
    Dim deleteCell As Range
    
    ' Get reference to the current sheet
    Set ws = ActiveSheet
    
    ' Add confirm box when deleting
    If ws.Name = "Template" Then
        MsgBox "Cannot delete the template"
        Exit Sub
    Else
        If MsgBox("Are you sure you want to perform this action?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Only works when all items are set to returned
    For Each c In Union(ws.Range("G6:G26"), ws.Range("G30:G34")).Cells
        If Not IsStringEmpty(c.Value) And Not c.Value = "Returned" Then
            MsgBox "There are still unreturned items."
            Exit Sub
        End If
    Next c
    
    ' Get the currentID from cell G2
    currentID = ws.Range("G2").Value
    
    ' Get reference to the "Menu" sheet
    On Error Resume Next
    Set menuSheet = ThisWorkbook.Sheets("Menu")
    On Error GoTo 0
    
    If Not menuSheet Is Nothing Then
        ' Get reference to the "MenuTable" as a ListObject (assumes it's a table)
        On Error Resume Next
        Set menuTable = menuSheet.ListObjects("MenuTable")
        On Error GoTo 0
        
        If Not menuTable Is Nothing Then
            ' Find the column in the table that matches the ID (assumes it's in column D)
            On Error Resume Next
            Set idColumn = menuTable.ListColumns("ID")
            On Error GoTo 0
            
            If Not idColumn Is Nothing Then
                ' Look for the currentID in the table's ID column
                On Error Resume Next
                Set deleteCell = idColumn.DataBodyRange.Find(What:=currentID, LookIn:=xlValues, LookAt:=xlWhole)
                On Error GoTo 0
                
                If Not deleteCell Is Nothing Then
                    ' Delete the entire row that contains the deleteCell
                    deleteCell.EntireRow.Delete
                Else
                    MsgBox "ID not found in MenuTable."
                End If
            Else
                MsgBox "Column 'ID' not found in MenuTable."
            End If
        Else
            MsgBox "Table 'MenuTable' not found in 'Menu' sheet."
        End If
    Else
        MsgBox "'Menu' sheet not found."
    End If
    
    
    ThisWorkbook.Sheets(Application.ActiveSheet.Name).Delete
    
End Sub