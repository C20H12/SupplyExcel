Sub ScanAllSheetsAndPrioritizeLabels()
    Dim SearchStrings As Variant
    Dim sh As Worksheet
    Dim wsMenu As Worksheet
    Dim pickupFound As Boolean
    Dim excludedSheetNames As Variant
    Dim barcode As String ' Move the barcode variable declaration outside the loop
    Dim foundStatus As String ' Move the foundStatus variable declaration outside the loop
    
    ' Define the search strings in the new order of priority
    SearchStrings = Array("S.O.S", "UNP", "Pick Up", "In Stock", "Ready To Order", "Ordered", "Complete", "Returned")
    
    ' Define the names of sheets to exclude from the search
    excludedSheetNames = Array("Menu", "Userform", "Template", "Pickup")
    
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
            barcode = sh.Cells(2, "G").Value ' Move this line here
            
            ' Initialize a string to store the found statuses
            foundStatus = "" ' Move this line here
            
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
                    ' If found, add the status to the foundStatus string
                    If foundStatus <> "" Then
                        foundStatus = foundStatus & ", "
                    End If
                    foundStatus = foundStatus & SearchString
                End If
            Next SearchString
            
            ' If "Ordered," "Ready To Order," or "Pick Up" statuses are found,
            ' update the Menu sheet with the concatenated statuses
            If foundStatus <> "" And (InStr(foundStatus, "Ordered") <> 0 Or InStr(foundStatus, "Ready To Order") <> 0 Or InStr(foundStatus, "Pick Up") <> 0) Then
                Dim menuCell As Range
                Set menuCell = wsMenu.Columns("E").Find(What:=barcode, LookIn:=xlValues, LookAt:=xlWhole)
                If Not menuCell Is Nothing Then
                    wsMenu.Cells(menuCell.Row, "C").Value = foundStatus
                    ' Set the pickupFound flag to True
                    pickupFound = True
                End If
            ElseIf foundStatus <> "" Then
                ' If other statuses are found (UNP, Complete, or Returned),
                ' prioritize and update the Menu sheet with the first found status
                Set menuCell = wsMenu.Columns("E").Find(What:=barcode, LookIn:=xlValues, LookAt:=xlWhole)
                If Not menuCell Is Nothing Then
                    wsMenu.Cells(menuCell.Row, "C").Value = Split(foundStatus, ", ")(0) ' Prioritize the first status
                    ' Set the pickupFound flag to True
                    pickupFound = True
                End If
            End If
        End If
    Next
    
    ' Check if no labels were found and display a message
    If Not pickupFound Then
        MsgBox "No labels found in any sheet (excluding 'Menu,' 'Userform,' and 'Template' sheets).", vbInformation
    End If
End Sub

Function IsInArray(val As Variant, arr As Variant) As Boolean
    Dim item As Variant
    For Each item In arr
        If item = val Then
            IsInArray = True
            Exit Function
        End If
    Next item
    IsInArray = False
End Function



