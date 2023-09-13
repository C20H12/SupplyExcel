Function ReCalculateSize()
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
        Dim SizeName As String
        SizeName = ActiveSheet.Range("B" & i).Value
        Dim ReturnedSize As String
        ReturnedSize = GetSize(SizeName, MeasuredSizes)
        Dim SplittedSize() As String
        SplittedSize = Split(ReturnedSize, "===")

        If Not Len(Trim(SizeName)) = 0 Then
            ActiveSheet.Range("E" & i).Value = SplittedSize(0)
            ActiveSheet.Range("A" & i).Value = SplittedSize(1)
        End If
    Next i
End Function