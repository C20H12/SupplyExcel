Sub Importing()
    
    
    ' loop over the used rows in this sheet
    For ii = 2 To ThisWorkbook.Sheets("Importing").UsedRange.Rows.count
    
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Sheets("Importing")
        
        ' Map all the variables needed to a cell in the row
        Dim FirstName As String
        Dim LastName As String
        Dim Rank As String
        Dim Email As String
        Dim Head As String
        Dim Neck As String
        Dim Chest As String
        Dim Waist As String
        Dim Hips As String
        Dim Height As String
        Dim FootL As String
        Dim FootW As String
        Dim Hand As String
        Dim Gender As Boolean
        
        FirstName = ws.Cells(ii, 3)
        LastName = ws.Cells(ii, 2)
        Rank = "AC"
        Email = ws.Cells(ii, 1)
        Gender = ws.Cells(ii, 4) = "Male"
        Head = ws.Cells(ii, 5)
        Neck = ws.Cells(ii, 6)
        Chest = ws.Cells(ii, 7)
        Waist = ws.Cells(ii, 8)
        Hips = ws.Cells(ii, 9)
        Height = ws.Cells(ii, 10)
        FootL = ws.Cells(ii, 11)
        FootW = ws.Cells(ii, 12)
        Hand = ws.Cells(ii, 13)
        
        ' Then do the exact same stuff as the NC form
        ' # Generate a ID for the new cadet and a sheet
        Dim sNewCadetID As String
        sNewCadetID = GetUUID()
        Dim sNewSheetName As String
        sNewSheetName = left(FirstName & "_" & LastName, 20) & "_" & sNewCadetID
        
        CreateNewCadetSheet (sNewSheetName)
        
            
        ' # Insert the values into the created sheet
        Sheets(sNewSheetName).Range("B2").Value = Rank
        Sheets(sNewSheetName).Range("C2").Value = LastName
        Sheets(sNewSheetName).Range("E2").Value = FirstName
        Sheets(sNewSheetName).Range("B4").Value = "9999999999"
        Sheets(sNewSheetName).Range("E4").Value = Email
        ' THIS IS SPECIFICALLY FOR THE REFERENCE CODE OF EACH CADET
        Sheets(sNewSheetName).Range("G2").Value = sNewCadetID
        
        If Gender = True Then
            Sheets(sNewSheetName).Range("G4").Value = "Female"
        Else
            Sheets(sNewSheetName).Range("G4").Value = "Male"
        End If
        
        Sheets(sNewSheetName).Range("L2").Value = Head
        Sheets(sNewSheetName).Range("L3").Value = Neck
        Sheets(sNewSheetName).Range("L4").Value = Chest
        Sheets(sNewSheetName).Range("L5").Value = Waist
        Sheets(sNewSheetName).Range("L6").Value = Hips
        Sheets(sNewSheetName).Range("L7").Value = Height
        Sheets(sNewSheetName).Range("L8").Value = FootL
        Sheets(sNewSheetName).Range("L9").Value = FootW
        Sheets(sNewSheetName).Range("L10").Value = Hand
        
        ' # Getting the sizing information
        Dim MeasuredSizes As Collection
        Set MeasuredSizes = New Collection
        MeasuredSizes.Add Head, "head"
        MeasuredSizes.Add Neck, "neck"
        MeasuredSizes.Add Chest, "chest"
        MeasuredSizes.Add Waist, "waist"
        MeasuredSizes.Add Hips, "hips"
        MeasuredSizes.Add Height, "height"
        MeasuredSizes.Add FootL, "FootL"
        MeasuredSizes.Add FootW, "FootW"
        MeasuredSizes.Add Hand, "hand"
        MeasuredSizes.Add Not Gender, "IsMale"
        
        For i = 6 To 24
            Dim SizeName As String
            SizeName = Sheets(sNewSheetName).Range("B" & i).Value
                    
            If Not IsStringEmpty(SizeName) Then
                Dim ReturnedSize As String
                ReturnedSize = GetSize(SizeName, MeasuredSizes)
                If Not IsStringEmpty(ReturnedSize) Then
                    Dim SplittedSize() As String
                    SplittedSize = Split(ReturnedSize, "===")
        
                    Sheets(sNewSheetName).Range("E" & i).Value = SplittedSize(0)
                    Sheets(sNewSheetName).Range("A" & i).Value = SplittedSize(1)
                End If
            End If
        Next i
        
        
        ' # Insert an entry to the menu that holds all sheets
        
        Set ws = ThisWorkbook.Sheets("Menu")
        
        ' Find the next empty row in column B of the "Menu" worksheet
        Dim empty_row As Long
        empty_row = ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 1
        
        ' Write the value from NC_FirstNameInput to the found empty row
        ' but using the self defined vars here
        ws.Cells(empty_row, 1).Value = LastName
        ws.Cells(empty_row, 2).Value = FirstName
        ws.Cells(empty_row, 4).Value = Now
        ws.Cells(empty_row, 5).Value = sNewCadetID
        
        Dim linkAddress As String
        linkAddress = "'" & sNewSheetName & "'!A1"
        
        ws.Hyperlinks.Add Anchor:=ws.Cells(empty_row, 1), _
                          Address:="", _
                          SubAddress:=linkAddress, _
                          TextToDisplay:=LastName
                          
            Columns("A:A").Select
        ActiveWorkbook.Worksheets("Menu").ListObjects("MenuTable").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Menu").ListObjects("MenuTable").Sort.SortFields. _
            Add Key:=Range("MenuTable[[#All],[Surname]]"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Menu").ListObjects("MenuTable").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next ii
End Sub

Sub ChangeFromOldColorsToNew()
    Dim ws As Worksheet
    Dim cell As Range

    ' Set the worksheet variable to the active sheet
    Set ws = ActiveSheet

    ' Loop through each cell in the worksheet
    For Each cell In ws.UsedRange
        ' Check if the cell's interior color is RGB(255, 153, 0)
        If cell.Interior.Color = RGB(0, 255, 0) Then
            ' Change the cell's interior color to RGB(244, 176, 132)
            cell.Interior.Color = RGB(251, 163, 251)
        ElseIf cell.Interior.Color = RGB(255, 153, 0) Then
            ' Change the cell's interior color to RGB(244, 176, 132)
            cell.Interior.Color = RGB(244, 176, 132)
        ElseIf cell.Interior.Color = RGB(74, 134, 232) Then
            ' Change the cell's interior color to RGB(244, 176, 132)
            cell.Interior.Color = RGB(155, 194, 230)
        ElseIf cell.Interior.Color = RGB(255, 0, 0) Then
            ' Change the cell's interior color to RGB(244, 176, 132)
            cell.Interior.Color = RGB(246, 246, 106)
        ElseIf cell.Interior.Color = RGB(0, 255, 255) Then
            ' Change the cell's interior color to RGB(244, 176, 132)
            cell.Interior.Color = RGB(146, 208, 80)
        ElseIf cell.Interior.Color = RGB(142, 124, 195) Then
            ' Change the cell's interior color to RGB(244, 176, 132)
            cell.Interior.Color = RGB(255, 255, 255)
        
        End If
    Next cell
End Sub
