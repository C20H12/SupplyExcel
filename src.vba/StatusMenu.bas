
Sub ScanAllSheetsAndPrioritizeLabels()
    Dim SearchStrings As Variant
    Dim sh As Worksheet
    Dim wsMenu As Worksheet
    Dim pickupFound As Boolean
    Dim excludedSheetNames As Variant
    
    ' Define the search strings in the new order of priority
    SearchStrings = Array("UNP", "Ready To Order", "Ordered", "Pick Up", "Complete")
    
    ' Define the names of sheets to exclude from the search
    excludedSheetNames = Array("Menu", "Userform", "Template")
    
    ' Set a reference to the Menu sheet
    On Error Resume Next
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    On Error GoTo 0
    
    If wsMenu Is Nothing Then
        MsgBox "Menu sheet not found in the workbook.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize the pickupFound flag
    pickupFound = False
    
    ' Loop through all sheets
    For Each sh In ActiveWorkbook.Worksheets
        ' Check if the sheet should be excluded
        If Not IsInArray(sh.Name, excludedSheetNames) Then
            ' Extract the unique barcode from cell G2
            Dim barcode As String
            barcode = sh.Cells(2, "G").Value
            
            ' Loop through each search string in the new order of priority
            For Each SearchString In SearchStrings
                ' Find the search string in the sheet
                Dim foundCell As Range
                Set foundCell = sh.Cells.Find(What:=SearchString, _
                    After:=sh.Cells(1, 1), _
                    LookIn:=xlValues, _
                    LookAt:=xlPart, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
                
                If Not foundCell Is Nothing Then
                    ' If found, update the Menu sheet with the label
                    Dim menuCell As Range
                    Set menuCell = wsMenu.Columns("E").Find(What:=barcode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not menuCell Is Nothing Then
                        wsMenu.Cells(menuCell.Row, "C").Value = SearchString
                        ' Set the pickupFound flag to True
                        pickupFound = True
                        Exit For ' Exit the loop once a label is found
                    End If
                End If
            Next SearchString
        End If
    Next
    
    ' Check if no labels were found and display a message
    If Not pickupFound Then
        MsgBox "No labels found in any sheet (excluding 'Menu,' 'Userform,' and 'Template' sheets).", vbInformation
    End If
End Sub

Function IsInArray(val As Variant, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = val Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function

