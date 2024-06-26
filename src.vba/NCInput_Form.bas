
Private Sub UserForm_Initialize()
'
' Initialize frm
' (This runs everytime the form is opened)
' frm is the form name
'
    ' Reset the size
    With NCInput_Form
        ' Set the form size
        Height = 620
        Width = 250
    End With
End Sub

Private Sub NC_CancelButton_Click()
    Unload Me
End Sub

Sub NC_SubmitButton_Click()
   Dim OrigBook As Workbook
   Set OrigBook = ActiveWorkbook
    ' so it does not trigger the cell change events, revert this setting after
    Application.EnableEvents = False
    
    ' # Validate all the input fields, skip if disabled
    Dim ValidateResults(1 To 23) As String
    
    ' Validate NC_FirstNameInput, NC_SurnameInput, NC_RankInput
    ValidateResults(1) = ValidateText(NC_FirstNameInput)
    ValidateResults(2) = ValidateText(NC_SurnameInput)
    ValidateResults(3) = ValidateText(NC_RankInput)

    
    ' Validate each size input to check if input is a number
    ValidateResults(6) = ValidateNumber(NC_HeadInput)
    ValidateResults(7) = ValidateNumber(NC_NeckInput)
    ValidateResults(8) = ValidateNumber(NC_ChestInput)
    ValidateResults(9) = ValidateNumber(NC_WaistInput)
    ValidateResults(10) = ValidateNumber(NC_HipsInput)
    ValidateResults(11) = ValidateNumber(NC_HeightInput)
    ValidateResults(12) = ValidateNumber(NC_FootLInput)
    ValidateResults(13) = ValidateNumber(NC_FootWInput)
    ValidateResults(14) = ValidateNumber(NC_HandLInput)
    
    ' Validate each input to check if size is in range
    If NC_EnableValidate Then
        ValidateResults(15) = ValidateRange(NC_HeadInput, 19, 24.6)
        ValidateResults(16) = ValidateRange(NC_NeckInput, 12.5, 20)
        ValidateResults(17) = ValidateRange(NC_ChestInput, 24, 53)
        ValidateResults(18) = ValidateRange(NC_WaistInput, 30, 63)
        ValidateResults(19) = ValidateRange(NC_HipsInput, 30, 68)
        ValidateResults(20) = ValidateRange(NC_HeightInput, 55, 76)
        ValidateResults(21) = ValidateRange(NC_FootLInput, 215, 330)
        ValidateResults(22) = ValidateRange(NC_FootWInput, 85, 130)
        ValidateResults(23) = ValidateRange(NC_HandLInput, 6, 10)
    End If

    ' Check if any validation fails, display error message (first encountered) and early return
    Dim ValidateTo As Integer
    ValidateTo = 14
    If NC_EnableValidate Then
        ValidateTo = 23
    End If
    For i = 1 To ValidateTo
        Dim ValidateResultMsg As String
        ValidateResultMsg = ValidateResults(i)
        If Not IsStringEmpty(ValidateResultMsg) Then
            MsgBox ValidateResultMsg, vbExclamation, "Input Error"
            Exit Sub
        End If
    Next i
    
    ' check if the user has either one of the telephone or email filled in
    If IsStringEmpty(NC_EmailInput.Value) And IsStringEmpty(NC_TelephoneInput) Then
        MsgBox "Must fill in at least one of: telephone or email!", vbExclamation, "Input Error"
        Exit Sub
    End If
        
    
    ' # Generate a ID for the new cadet and a sheet
    Dim sNewCadetID As String
    sNewCadetID = GetUUID()
    Dim sNewSheetName As String
    sNewSheetName = left(Replace(NC_FirstNameInput.Value, " ", "_") & "_" & Replace(NC_SurnameInput.Value, " ", "_"), 20) & "_" & sNewCadetID

    CreateNewCadetSheet (sNewSheetName)
  
    Dim CreatedSheet As Worksheet
    Set CreatedSheet = Sheets(sNewSheetName)
        
    ' # Insert the values into the created sheet
    CreatedSheet.Range("B2").Value = NC_RankInput.Value
    CreatedSheet.Range("C2").Value = NC_SurnameInput.Value
    CreatedSheet.Range("E2").Value = NC_FirstNameInput.Value
    CreatedSheet.Range("B4").Value = NC_TelephoneInput.Value
    CreatedSheet.Range("E4").Value = NC_EmailInput.Value
    ' THIS IS SPECIFICALLY FOR THE REFERENCE CODE OF EACH CADET
    CreatedSheet.Range("G2").Value = sNewCadetID
    
    If NC_FemaleInput = True Then
        CreatedSheet.Range("G4").Value = "Female"
    Else
        CreatedSheet.Range("G4").Value = "Male"
    End If
    
    CreatedSheet.Range("L2").Value = NC_HeadInput.Value
    CreatedSheet.Range("L3").Value = NC_NeckInput.Value
    CreatedSheet.Range("L4").Value = NC_ChestInput.Value
    CreatedSheet.Range("L5").Value = NC_WaistInput.Value
    CreatedSheet.Range("L6").Value = NC_HipsInput.Value
    CreatedSheet.Range("L7").Value = NC_HeightInput.Value
    CreatedSheet.Range("L8").Value = NC_FootLInput.Value
    CreatedSheet.Range("L9").Value = NC_FootWInput.Value
    CreatedSheet.Range("L10").Value = NC_HandLInput.Value
    
    ' # Getting the sizing information
    Dim MeasuredSizes As Collection
    Set MeasuredSizes = New Collection
    MeasuredSizes.Add NC_HeadInput.Value, "head"
    MeasuredSizes.Add NC_NeckInput.Value, "neck"
    MeasuredSizes.Add NC_ChestInput.Value, "chest"
    MeasuredSizes.Add NC_WaistInput.Value, "waist"
    MeasuredSizes.Add NC_HipsInput.Value, "hips"
    MeasuredSizes.Add NC_HeightInput.Value, "height"
    MeasuredSizes.Add NC_FootLInput.Value, "FootL"
    MeasuredSizes.Add NC_FootWInput.Value, "FootW"
    MeasuredSizes.Add NC_HandLInput.Value, "hand"
    MeasuredSizes.Add Not NC_FemaleInput, "IsMale"
    
    For i = 6 To 24
        Dim sizeName As String
        sizeName = CreatedSheet.Range("B" & i).Value
                
        If Not IsStringEmpty(sizeName) Then
            Dim ReturnedSize As String
            ReturnedSize = GetSize(sizeName, MeasuredSizes)
            If Not IsStringEmpty(ReturnedSize) Then
                Dim SplittedSize() As String
                SplittedSize = Split(ReturnedSize, "===")
               ' Debug.Print ReturnedSize
               ' Debug.Print ReturnedSize, SplittedSize(0)
    
                CreatedSheet.Range("E" & i).Value = SplittedSize(0)
                CreatedSheet.Range("A" & i).Value = SplittedSize(1)
                
                If NC_EnableStock And FindInInventory(SplittedSize(1)) > 0 Then
                    CreatedSheet.Range("G" & i).Value = "In Stock"
                End If
            End If
        End If
    Next i
    
    ' # Insert an entry to the menu that holds all sheets
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Menu")
    
    ' Find the next empty row in column B of the "Menu" worksheet
    Dim empty_row As Long
    empty_row = ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 1
    
    ' Write the value from NC_FirstNameInput to the found empty row
    ws.Cells(empty_row, 1).Value = NCInput_Form.NC_SurnameInput.Value
    ws.Cells(empty_row, 2).Value = NCInput_Form.NC_FirstNameInput.Value
    ws.Cells(empty_row, 4).Value = Now
    ws.Cells(empty_row, 5).Value = sNewCadetID
    
    Dim linkAddress As String
    linkAddress = "'" & sNewSheetName & "'!A1"
    
    ws.Hyperlinks.Add Anchor:=ws.Cells(empty_row, 1), _
                      Address:="", _
                      SubAddress:=linkAddress, _
                      TextToDisplay:=NC_SurnameInput.Value
                      
        Columns("A:A").Select
    OrigBook.Worksheets("Menu").ListObjects("MenuTable").Sort.SortFields. _
        Clear
    OrigBook.Worksheets("Menu").ListObjects("MenuTable").Sort.SortFields. _
        Add Key:=OrigBook.Worksheets("Menu").Range("MenuTable[[#All],[Surname]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With OrigBook.Worksheets("Menu").ListObjects("MenuTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.EnableEvents = True
    
    Unload Me
    
End Sub


Private Sub UserForm_Click()

End Sub