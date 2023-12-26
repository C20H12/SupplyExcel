Sub Resize()
    ' confirm box
    If MsgBox("Are you sure you want to perform this action?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    ReCalculateSize
    
    Application.EnableEvents = True
End Sub
Sub ReCalculateSize(Optional ItemNameToChange As String)
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
            
            ' if nothing is passed in, do sizing OR if an exact item name is passed in, it needs to match
            If IsStringEmpty(ItemNameToChange) Or (Not IsStringEmpty(ItemNameToChange) And ItemNameToChange = ItemName) Then
                Dim ReturnedSize As String
                ReturnedSize = GetSize(ItemName, MeasuredSizes)
                
                If Not IsStringEmpty(ReturnedSize) Then
                    Dim SplittedSize() As String
                    SplittedSize = Split(ReturnedSize, "===")
                    
                    ' if size has changed, change status to unp
                    ' Use .Text so that fractions aren't converted to decimals
                    If Not SplittedSize(0) = ActiveSheet.Range("E" & i).Text Then
                        ActiveSheet.Range("G" & i).Value = "UNP"
                    End If
    
                    ActiveSheet.Range("E" & i).Value = SplittedSize(0)
                    ActiveSheet.Range("A" & i).Value = SplittedSize(1)
                End If
            End If
        End If
    Next i
End Sub
