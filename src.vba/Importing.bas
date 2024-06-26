
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


Private Function GetStatusFromColorCell(cell As Range) As String
    If cell.Interior.Color = RGB(255, 117, 117) Then
      GetStatusFromColorCell = "UNP"
    ElseIf cell.Interior.Color = RGB(251, 163, 251) Then
      GetStatusFromColorCell = "In Stock"
    ElseIf cell.Interior.Color = RGB(146, 208, 80) Then
      GetStatusFromColorCell = "Pick Up"
    ElseIf cell.Interior.Color = RGB(246, 246, 106) Then
      GetStatusFromColorCell = "Ready To Order"
    ElseIf cell.Interior.Color = RGB(244, 176, 132) Then
      GetStatusFromColorCell = "Ordered"
    ElseIf cell.Interior.Color = RGB(155, 194, 230) Then
      GetStatusFromColorCell = "Complete"
    ElseIf cell.Interior.Color = RGB(128, 128, 128) Then
      GetStatusFromColorCell = "Returned"
    Else
      GetStatusFromColorCell = "Unknown"
    End If
End Function


Sub ChangeFraction()
Set ws = ActiveWorkbook.Sheets("Import Sheets")
Dim i As Integer
Dim Fraction As String
lastRow = ws.UsedRange.Rows.count

For i = 2 To lastRow
    Number = ws.Cells(i, 16).Value
    If Not Number = Int(Number) Then
        Fraction = DecimalToComplexFraction(Number)
        ws.Cells(i, 16).Value = Fraction
    End If
Next i

For i = 2 To lastRow
    Number = ws.Cells(i, 24).Value
    If Not Number = Int(Number) Then
        Fraction = DecimalToComplexFraction(Number)
        ws.Cells(i, 24).Value = Fraction
    End If
Next i
    
End Sub
Function DecimalToComplexFraction(ByVal decimalNumber As String) As String
    Dim wholePart As Integer
    Dim numerator As Integer
    Dim denominator As Integer
    Dim tolerance As Double

    tolerance = 0
    
    wholePart = Int(decimalNumber)
    decimalNumber = decimalNumber - wholePart
    numerator = 1
    denominator = 1
    
    Do While Abs(numerator / denominator - decimalNumber) > tolerance
        If numerator / denominator < decimalNumber Then
            numerator = numerator + 1
        Else
            denominator = denominator + 1
        End If
    Loop
    
    DecimalToComplexFraction = wholePart & " " & Format(numerator, "#") & "/" & Format(denominator, "#")
End Function


Sub ImportFromOldSheet()
    Application.EnableEvents = False
    
    ' loop over the used rows in this sheet
    For Row = 2 To ThisWorkbook.Sheets("Import Sheets").UsedRange.Rows.count
    
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Sheets("Import Sheets")
        
        ' Map all the variables needed to a cell in the row
        Dim LastName As String
        Dim FirstName As String
        Dim Rank As String
        Dim Gender As Boolean
        Dim Head As String
        Dim Neck As String
        Dim Chest As String
        Dim Waist As String
        Dim Hips As String
        Dim Height As String
        Dim FootL As String
        Dim FootW As String
        Dim Hand As String
        Dim ID As String
        Dim Email As String
        Dim Phone As String
        Dim shouldGenerate As Boolean

        Dim sizes(6 To 26) As Range

        LastName = ws.Cells(Row, 1)
        FirstName = ws.Cells(Row, 2)
        Gender = ws.Cells(Row, 3)
        Rank = IIf(IsStringEmpty(ws.Cells(Row, 4)), "AC", ws.Cells(Row, 4))
        ID = ws.Cells(Row, 5)
        Head = ws.Cells(Row, 6)
        Neck = ws.Cells(Row, 7)
        Chest = ws.Cells(Row, 8)
        Waist = ws.Cells(Row, 9)
        Hips = ws.Cells(Row, 10)
        Height = ws.Cells(Row, 11)
        FootL = ws.Cells(Row, 12)
        FootW = ws.Cells(Row, 13)
        Hand = ws.Cells(Row, 14)
        Phone = ws.Cells(Row, 33)
        Email = ws.Cells(Row, 34)
        shouldGenerate = ws.Cells(Row, 36) = "Y"
        
        Set sizes(6) = ws.Cells(Row, 15)  ' Tunic
        Set sizes(7) = ws.Cells(Row, 17) ' Shirt
        Set sizes(8) = ws.Cells(Row, 18)  ' TShirt
        Set sizes(9) = ws.Cells(Row, 16)  ' Pants
        Set sizes(10) = ws.Cells(Row, 19)  ' Wedge
        Set sizes(11) = ws.Cells(Row, 20)  ' Tie
        Set sizes(12) = ws.Cells(Row, 21)  ' PantBelt
        Set sizes(13) = ws.Cells(Row, 22)  ' Socks
        Set sizes(14) = ws.Cells(Row, 23)  ' Boots

        Set sizes(16) = ws.Cells(Row, 30)  ' Toque
        Set sizes(17) = ws.Cells(Row, 31)  ' Tilly
        Set sizes(18) = ws.Cells(Row, 28)  ' Parka
        Set sizes(19) = ws.Cells(Row, 29)  ' Gloves

        Set sizes(21) = ws.Cells(Row, 27)  ' Beret
        Set sizes(22) = ws.Cells(Row, 24)  ' FTUShirt
        Set sizes(23) = ws.Cells(Row, 25)  ' FTUPants
        Set sizes(24) = ws.Cells(Row, 26)  ' FTUBoots
        Set sizes(26) = ws.Cells(Row, 32)  'NameTag
        
         ' Then do the exact same stuff as the NC form
        ' # Generate a ID for the new cadet and a sheet
        Dim sNewCadetID As String
        ' if ID is empty, generate one
        sNewCadetID = IIf(IsStringEmpty(ID), GetUUID(), ID)
        Dim sNewSheetName As String
        sNewSheetName = left(Replace(FirstName, " ", "_") & "_" & Replace(LastName, " ", "_"), 20) & "_" & sNewCadetID
        
        If SheetExist(sNewSheetName) Then
            GoTo continue
        End If
        
        CreateNewCadetSheet (sNewSheetName)
        
        ' # Insert the values into the created sheet
        Sheets(sNewSheetName).Range("B2").Value = Rank
        Sheets(sNewSheetName).Range("C2").Value = LastName
        Sheets(sNewSheetName).Range("E2").Value = FirstName
        Sheets(sNewSheetName).Range("B4").Value = Phone
        Sheets(sNewSheetName).Range("E4").Value = Email
        ' THIS IS SPECIFICALLY FOR THE REFERENCE CODE OF EACH CADET
        Sheets(sNewSheetName).Range("G2").Value = sNewCadetID
        
        
        Sheets(sNewSheetName).Range("L2").Value = Head
        Sheets(sNewSheetName).Range("L3").Value = Neck
        Sheets(sNewSheetName).Range("L4").Value = Chest
        Sheets(sNewSheetName).Range("L5").Value = Waist
        Sheets(sNewSheetName).Range("L6").Value = Hips
        Sheets(sNewSheetName).Range("L7").Value = Height
        Sheets(sNewSheetName).Range("L8").Value = FootL
        Sheets(sNewSheetName).Range("L9").Value = FootW
        Sheets(sNewSheetName).Range("L10").Value = Hand
        
        If shouldGenerate Then
            ' # Getting the sizing information
            Dim MeasuredSizes As Collection
            Set MeasuredSizes = New Collection
            MeasuredSizes.Add ActiveSheet.Range("L2").Value, "head"
            MeasuredSizes.Add ActiveSheet.Range("L3").Value, "neck"
            MeasuredSizes.Add ActiveSheet.Range("L4").Value, "chest"
            MeasuredSizes.Add ActiveSheet.Range("L5").Value, "waist"
            MeasuredSizes.Add ActiveSheet.Range("L6").Value, "hips"
            MeasuredSizes.Add ActiveSheet.Range("L7").Value, "height"
            MeasuredSizes.Add ActiveSheet.Range("L8").Value, "FootL"
            MeasuredSizes.Add ActiveSheet.Range("L9").Value, "FootW"
            MeasuredSizes.Add ActiveSheet.Range("L10").Value, "hand"
            MeasuredSizes.Add ActiveSheet.Range("G4").Value = "Male", "IsMale"
            
            For i = 6 To 24
                Dim ItemName As String
                ItemName = ActiveSheet.Range("B" & i).Value
                
                ' only check non empty cells in the item names column
                If Not IsStringEmpty(ItemName) Then
                    Dim ReturnedSize As String
                    ReturnedSize = GetSize(ItemName, MeasuredSizes)
                    
                    If Not IsStringEmpty(ReturnedSize) Then
                        Dim SplittedSize() As String
                        SplittedSize = Split(ReturnedSize, "===")
                
                        ActiveSheet.Range("E" & i).Value = SplittedSize(0)
                        ActiveSheet.Range("A" & i).Value = SplittedSize(1)
                    End If
                End If
            Next i
        Else
            For i = 6 To 24
                Dim sizeName As String
                sizeName = Sheets(sNewSheetName).Range("B" & i).Value
                        
                If Not IsStringEmpty(sizeName) Then
                    Dim SizeNSN As String
                    SizeNSN = GetNSNFromSize(sizeName, sizes(i).Value, Gender)
                    If IsStringEmpty(SizeNSN) And (sizeName = "Dress Pants" Or sizeName = "Collar Shirt") Then
                        SizeNSN = GetNSNFromSize(sizeName, sizes(i).Value, Not Gender)
                    End If
                    If Not IsStringEmpty(SizeNSN) Then
                        Sheets(sNewSheetName).Range("A" & i).Value = SizeNSN
                    End If
                    Sheets(sNewSheetName).Range("E" & i).Value = sizes(i).Value
                    Sheets(sNewSheetName).Range("G" & i).Value = GetStatusFromColorCell(sizes(i))
                End If
            Next i
        End If
        
        ' nametag
        Sheets(sNewSheetName).Range("G" & 26).Value = GetStatusFromColorCell(sizes(26))
        
        ' exchange history
        If Not IsStringEmpty(ws.Cells(Row, 35)) Then
            Dim exchangeRows() As String
            exchangeRows = Split(ws.Cells(Row, 35).Value, "+++")
            For Each exchangerow In exchangeRows
                Dim exchangeRowItemList()  As String
                exchangeRowItemList = Split(exchangerow, "===")
                Dim exchangeRowTable As ListRow
                Set exchangeRowTable = Sheets(sNewSheetName).ListObjects(sNewSheetName & "ExchangeTable").ListRows.Add
                exchangeRowTable.Range(1, 1).Value = exchangeRowItemList(0)
                exchangeRowTable.Range(1, 2).Value = exchangeRowItemList(1)
                exchangeRowTable.Range(1, 3).Value = exchangeRowItemList(2)
                exchangeRowTable.Range(1, 4).Value = exchangeRowItemList(3)
            Next exchangerow
        End If
        
        
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
continue:
    Next Row

For i = 2 To Numinputs
    ws.Rows(2).Delete
Next i

    Application.EnableEvents = True
End Sub


Sub ExportData()
    Application.EnableEvents = False
    
    Dim OrigBook As Workbook, OutBook As Workbook
    Set OrigBook = ThisWorkbook
    Set OutBook = Workbooks.Add
    formattedDateTime = Format(Now, "hh-mm-ss")
    OutBook.SaveAs OrigBook.Path & Application.PathSeparator & "Exported_Data at " & formattedDateTime & ".xlsx"
    
    Dim ows As Worksheet
    Set ows = OutBook.Sheets("Sheet1")
    Dim ws As Worksheet
    
    Dim Row As Integer
    Row = 0
    
    For i = 1 To 34
        ows.Columns(i).NumberFormat = "@"
    Next i
    
    For Each ws In OrigBook.Sheets
        If isSpecialSheet(ws.name) Then
            GoTo continue
        End If
        
        ' Map all the variables needed to a cell in the row
        Row = Row + 1
        
        ows.Cells(Row, 1) = ws.Range("C2").Value    ' User Last Name
        ows.Cells(Row, 2) = ws.Range("E2").Value    ' User First Name
        ows.Cells(Row, 3) = ws.Range("G4").Value    ' User Gender
        ows.Cells(Row, 4) = ws.Range("B2").Value   ' User Rank
        ows.Cells(Row, 5) = ws.Range("G2").Value   ' User ID
        ows.Cells(Row, 6) = ws.Range("L2").Value    ' User Head
        ows.Cells(Row, 7) = ws.Range("L3").Value    ' User Neck
        ows.Cells(Row, 8) = ws.Range("L4").Value    ' User Chest
        ows.Cells(Row, 9) = ws.Range("L5").Value    ' User Waist
        ows.Cells(Row, 10) = ws.Range("L6").Value    ' User Hips
        ows.Cells(Row, 11) = ws.Range("L7").Value    ' User Height
        ows.Cells(Row, 12) = ws.Range("L8").Value   ' User Foot Length
        ows.Cells(Row, 13) = ws.Range("L9").Value   ' User Foot Width
        ows.Cells(Row, 14) = ws.Range("L10").Value   ' User Hand
        
        

        
        
        ows.Cells(Row, 15) = ws.Cells(6, 5)  ' Tunic
        ows.Cells(Row, 15).Interior.Color = ws.Cells(6, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 16) = ws.Cells(9, 5) ' Pants
        ows.Cells(Row, 16).Interior.Color = ws.Cells(9, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 17) = ws.Cells(7, 5) 'Shirts
        ows.Cells(Row, 17).Interior.Color = ws.Cells(7, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 18) = ws.Cells(8, 5)   ' TShirt
        ows.Cells(Row, 18).Interior.Color = ws.Cells(8, 7).DisplayFormat.Interior.Color
        
       
        ows.Cells(Row, 19) = ws.Cells(10, 5)  ' Wedge
        ows.Cells(Row, 19).Interior.Color = ws.Cells(10, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 20) = ws.Cells(11, 5)  ' Tie
        ows.Cells(Row, 20).Interior.Color = ws.Cells(11, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 21) = ws.Cells(12, 5)  ' PantBelt
        ows.Cells(Row, 21).Interior.Color = ws.Cells(12, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 22) = ws.Cells(13, 5)  ' Socks
        ows.Cells(Row, 22).Interior.Color = ws.Cells(13, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 23) = ws.Cells(14, 5)  ' Boots
        ows.Cells(Row, 23).Interior.Color = ws.Cells(14, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 24) = ws.Cells(22, 5)  ' FTUShirt
        ows.Cells(Row, 24).Interior.Color = ws.Cells(22, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 25) = ws.Cells(23, 5)  ' FTUPants
        ows.Cells(Row, 25).Interior.Color = ws.Cells(23, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 26) = ws.Cells(24, 5)  ' FTUBoots
        ows.Cells(Row, 26).Interior.Color = ws.Cells(24, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 27) = ws.Cells(21, 5)  ' Beret
        ows.Cells(Row, 27).Interior.Color = ws.Cells(26, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 28) = ws.Cells(18, 5)  ' Parka
        ows.Cells(Row, 28).Interior.Color = ws.Cells(17, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 29) = ws.Cells(19, 5)  ' Gloves
        ows.Cells(Row, 29).Interior.Color = ws.Cells(18, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 30) = ws.Cells(16, 5)  ' Toque
        ows.Cells(Row, 30).Interior.Color = ws.Cells(16, 7).DisplayFormat.Interior.Color
         

        ows.Cells(Row, 31) = ws.Cells(17, 5)  ' Tilly
        ows.Cells(Row, 31).Interior.Color = ws.Cells(17, 7).DisplayFormat.Interior.Color
        
        ows.Cells(Row, 32) = ws.Cells(26, 5)  ' Nametag
        ows.Cells(Row, 32).Interior.Color = ws.Cells(26, 7).DisplayFormat.Interior.Color
        
        
        ows.Cells(Row, 33) = "'" & ws.Cells(4, 2).Value  ' phone

        ows.Cells(Row, 34) = ws.Cells(4, 5).Value   ' email
        
        ' exchange history as string date===item===size===size_new+++[more]
        If Not IsStringEmpty(ws.Cells(38, 1).Value) Then
            Dim exchangeHistory() As String
            Dim exchangeHistoryRows As Integer
            exchangeHistoryRows = ws.ListObjects(ws.name & "ExchangeTable").Range.Rows.count
            
            For i = 0 To exchangeHistoryRows - 2
                Dim exchangeHistoryRow() As Variant
                exchangeHistoryRow = Array(ws.Cells(38 + i, 1), ws.Cells(38 + i, 2), ws.Cells(38 + i, 3), ws.Cells(38 + i, 4))
                ReDim Preserve exchangeHistory(i)
                exchangeHistory(i) = Join(exchangeHistoryRow, "===")
            Next i
            
            ows.Cells(Row, 35) = Join(exchangeHistory, "+++")
       
        End If
continue:
    Next ws


    OutBook.Close True
    Application.EnableEvents = True
End Sub




