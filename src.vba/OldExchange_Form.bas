
Option Explicit

Dim HeadCount As Integer
Dim NeckCount As Integer
Dim ChestCount As Integer
Dim WaistCount As Integer
Dim HipsCount As Integer
Dim HeightCount As Integer
Dim FootLCount As Integer
Dim FootWCount As Integer
Dim HandCount As Integer
Dim i As Integer

Private Sub UserForm_Initialize()
'
' Initialize frm
' (This runs everytime the form is opened)
' frm is the form name
'
    ' Reset the size
    With OldExchange_Form
        ' Set the form size
        Height = 450
        Width = 500
    End With
    
'    EX_TShirtToggle = True
'    EX_FirstNameInput = "john"
'    EX_SurnameInput = "ignoroff"
'    EX_RankInput = "AC"
'    EX_HeadInput.Value = GenerateRandomMeasurement(19, 26)
'    EX_NeckInput.Value = GenerateRandomMeasurement(12.5, 20)
'    EX_ChestInput.Value = GenerateRandomMeasurement(24, 64)
'    EX_WaistInput.Value = GenerateRandomMeasurement(30, 63)
'    EX_HipsInput.Value = GenerateRandomMeasurement(30, 68)
'    EX_HeightInput.Value = GenerateRandomMeasurement(55, 76)
'    EX_FootLInput.Value = GenerateRandomMeasurement(215, 330)
'    EX_FootWInput.Value = GenerateRandomMeasurement(85, 130)
'    EX_HandLInput.Value = GenerateRandomMeasurement(6, 10)
'    EX_FemaleInput.Value = GenerateRandomFemale()
End Sub

' # Form controls
Private Sub EX_CancelButton_Click()

Unload Me

End Sub

Private Sub EX_SubmitButton_Click()

    Application.EnableEvents = False
    
    If HeadCount = 0 And _
        NeckCount = 0 And _
        ChestCount = 0 And _
        WaistCount = 0 And _
        HipsCount = 0 And _
        HeightCount = 0 And _
        FootLCount = 0 And _
        FootWCount = 0 And _
        HandCount = 0 Then
        
        MsgBox "Please select an item to exchange", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    Dim ValidateResults(1 To 3) As String
    ' Validate EX_FirstNameInput, EX_SurnameInput, EX_RankInput
    ValidateResults(1) = ValidateText(EX_FirstNameInput)
    ValidateResults(2) = ValidateText(EX_SurnameInput)
    ValidateResults(3) = ValidateText(EX_RankInput)
    
    For i = 1 To 3
        Dim ValidateResultMsg As String
        ValidateResultMsg = ValidateResults(i)
        If Not IsStringEmpty(ValidateResultMsg) Then
            MsgBox ValidateResultMsg, vbExclamation, "Input Error"
            Exit Sub
        End If
    Next i




    ' do validation on each
    Dim bPassed As Boolean
    If EX_EnableValidate Then
        bPassed = EX_DataValidation
    Else
        bPassed = True
    End If
    If Not bPassed Then
        Exit Sub
    End If
    
    ' confirm box
    If MsgBox("Are you sure you want to perform this action?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    'Create Array selectbuttons
    Dim SelectButtons(1 To 16) As MSForms.ToggleButton
    Set SelectButtons(1) = EX_TunicToggle
    Set SelectButtons(2) = EX_CollaredShirtToggle
    Set SelectButtons(3) = EX_TShirtToggle
    Set SelectButtons(4) = EX_DressPantsToggle
    Set SelectButtons(5) = EX_WedgeToggle
    Set SelectButtons(6) = EX_TieToggle
    Set SelectButtons(7) = EX_BeltToggle
    Set SelectButtons(8) = EX_SocksToggle
    Set SelectButtons(9) = EX_LeatherBootsToggle
    Set SelectButtons(10) = EX_TillyToggle
    Set SelectButtons(11) = EX_ParkaToggle
    Set SelectButtons(12) = EX_GlovesToggle
    Set SelectButtons(13) = EX_BeretToggle
    Set SelectButtons(14) = EX_FTUTunicToggle
    Set SelectButtons(15) = EX_FTUPantsToggle
    Set SelectButtons(16) = EX_FTUBootsToggle
    
    
    
    Dim sNewCadetID As String
    sNewCadetID = GetUUID()
    Dim sNewSheetName As String
    sNewSheetName = left(EX_FirstNameInput.Value & "_" & EX_SurnameInput.Value, 20) & "_" & sNewCadetID

    
    CreateNewCadetSheet (sNewSheetName)
    
    Sheets(sNewSheetName).Range("B2").Value = EX_RankInput.Value
    Sheets(sNewSheetName).Range("C2").Value = EX_SurnameInput.Value
    Sheets(sNewSheetName).Range("E2").Value = EX_FirstNameInput.Value
    Sheets(sNewSheetName).Range("B4").Value = EX_TelephoneInput.Value
    Sheets(sNewSheetName).Range("E4").Value = EX_EmailInput.Value
    ' THIS IS SPECIFICALLY FOR THE REFERENCE CODE OF EACH CADET
    Sheets(sNewSheetName).Range("G2").Value = sNewCadetID
    
    If EX_FemaleInput = True Then
        Sheets(sNewSheetName).Range("G4").Value = "Female"
    Else
        Sheets(sNewSheetName).Range("G4").Value = "Male"
    End If
    
    Sheets(sNewSheetName).Range("L2").Value = EX_HeadInput.Value
    Sheets(sNewSheetName).Range("L3").Value = EX_NeckInput.Value
    Sheets(sNewSheetName).Range("L4").Value = EX_ChestInput.Value
    Sheets(sNewSheetName).Range("L5").Value = EX_WaistInput.Value
    Sheets(sNewSheetName).Range("L6").Value = EX_HipsInput.Value
    Sheets(sNewSheetName).Range("L7").Value = EX_HeightInput.Value
    Sheets(sNewSheetName).Range("L8").Value = EX_FootLInput.Value
    Sheets(sNewSheetName).Range("L9").Value = EX_FootWInput.Value
    Sheets(sNewSheetName).Range("L10").Value = EX_HandLInput.Value
    
    ' # Getting the sizing information
    Dim MeasuredSizes As Collection
    Set MeasuredSizes = New Collection
    MeasuredSizes.Add IIf(IsStringEmpty(EX_HeadInput.Value), 0, EX_HeadInput.Value), "head"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_NeckInput.Value), 0, EX_NeckInput.Value), "neck"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_ChestInput.Value), 0, EX_ChestInput.Value), "chest"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_WaistInput.Value), 0, EX_WaistInput.Value), "waist"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_HipsInput.Value), 0, EX_HipsInput.Value), "hips"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_HeightInput.Value), 0, EX_HeightInput.Value), "height"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_FootLInput.Value), 0, EX_FootLInput.Value), "FootL"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_FootWInput.Value), 0, EX_FootWInput.Value), "FootW"
    MeasuredSizes.Add IIf(IsStringEmpty(EX_HandLInput.Value), 0, EX_HandLInput.Value), "hand"
    MeasuredSizes.Add Not EX_FemaleInput, "IsMale"
    
    For i = 6 To 24
        Dim SizeName As String
        SizeName = Sheets(sNewSheetName).Range("B" & i).Value
                
        If Not IsStringEmpty(SizeName) Then
            Dim ReturnedSize As String
            ReturnedSize = GetSize(SizeName, MeasuredSizes)
            If Not IsStringEmpty(ReturnedSize) Then
                Dim SplittedSize() As String
                SplittedSize = Split(ReturnedSize, "===")
               ' Debug.Print ReturnedSize
               ' Debug.Print ReturnedSize, SplittedSize(0)
    
                Sheets(sNewSheetName).Range("E" & i).Value = SplittedSize(0)
                Sheets(sNewSheetName).Range("A" & i).Value = SplittedSize(1)
            End If
        End If
    Next i
    
    ' # Insert an entry to the menu that holds all sheets
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Menu")
    
    ' Find the next empty row in column B of the "Menu" worksheet
    Dim empty_row As Long
    empty_row = ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 1
    
    ' Write the value from EX_FirstNameInput to the found empty row
    ws.Cells(empty_row, 1).Value = OldExchange_Form.EX_SurnameInput.Value
    ws.Cells(empty_row, 2).Value = OldExchange_Form.EX_FirstNameInput.Value
    ws.Cells(empty_row, 4).Value = Now
    ws.Cells(empty_row, 5).Value = sNewCadetID
    
    Dim linkAddress As String
    linkAddress = "'" & sNewSheetName & "'!A1"
    
    ws.Hyperlinks.Add Anchor:=ws.Cells(empty_row, 1), _
                      Address:="", _
                      SubAddress:=linkAddress, _
                      TextToDisplay:=EX_SurnameInput.Value
                      
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
    
    
    Dim nws As Worksheet
    Set nws = ActiveWorkbook.Worksheets(sNewSheetName)
    Dim extbl As ListObject
    Set extbl = nws.ListObjects(sNewSheetName & "ExchangeTable")

    Dim SelectedButton As Variant
    Dim SelectedItems() As String
    Dim numexch As Integer
    
    numexch = 0
    For Each SelectedButton In SelectButtons
        If SelectedButton Then
         numexch = numexch + 1
         ReDim Preserve SelectedItems(1 To numexch)
         SelectedItems(numexch) = CStr(SelectedButton.Caption)
        End If
    Next SelectedButton
    
  
    For i = 6 To 26
        If Not IsInArray(nws.Range("B" & CStr(i)).Value, SelectedItems()) Then
            ' Clear the value in column E for the current row
            nws.Range("E" & CStr(i)).Value = "---------"
            nws.Range("A" & CStr(i)).Value = ""
            nws.Range("G" & CStr(i)).Value = "Complete"
        Else
            ' Add a new row to the ExchangeTable
            Dim NewRow As ListRow
            Set NewRow = extbl.ListRows.Add
            NewRow.Range.Cells(1, 1) = Format(Date, "yyyy-mm-dd")
            NewRow.Range.Cells(1, 2) = nws.Range("B" & CStr(i)).Value
            NewRow.Range.Cells(1, 3) = InputBox("Previous " & nws.Range("B" & CStr(i)).Value & " Size", "Exchange Data")
            NewRow.Range.Cells(1, 4) = nws.Range("E" & CStr(i)).Value
        End If
    Next i

    
    
    Application.EnableEvents = True
    
    Unload Me

End Sub

' # Toggle Buttons
Private Sub EX_GlovesToggle_Click()
    If EX_GlovesToggle.Value = True Then
        HandCount = HandCount + 1
    ElseIf EX_LeatherBootsToggle.Value = False Then
        HandCount = HandCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_LeatherBootsToggle_Click()
  
  If EX_LeatherBootsToggle.Value = True Then
        FootLCount = FootLCount + 1
        FootWCount = FootWCount + 1
    ElseIf EX_LeatherBootsToggle.Value = False Then
        FootLCount = FootLCount - 1
        FootWCount = FootWCount - 1
    End If
    UpdateCounts
End Sub


Private Sub EX_FTUTunicToggle_Click()
    If EX_FTUTunicToggle.Value = True Then
        ChestCount = ChestCount + 1
        HeightCount = HeightCount + 1
    ElseIf EX_FTUTunicToggle.Value = False Then
        ChestCount = ChestCount - 1
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_FTUPantsToggle_Click()
    If EX_FTUPantsToggle.Value = True Then
        WaistCount = WaistCount + 1
        HeightCount = HeightCount + 1
    ElseIf EX_FTUPantsToggle.Value = False Then
        WaistCount = WaistCount - 1
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub
Private Sub EX_FTUBootsToggle_Click()
    If EX_FTUBootsToggle.Value = True Then
        FootLCount = FootLCount + 1
        FootWCount = FootWCount + 1
    ElseIf EX_FTUBootsToggle.Value = False Then
        FootLCount = FootLCount - 1
        FootWCount = FootWCount - 1
    End If
    UpdateCounts
End Sub
Private Sub EX_SocksToggle_Click()
    If EX_SocksToggle.Value = True Then
        FootLCount = FootLCount + 1
    ElseIf EX_FTUBootsToggle.Value = False Then
        FootLCount = FootLCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_TieToggle_Click()
    If EX_TieToggle.Value = True Then
        NeckCount = NeckCount + 1
    ElseIf EX_TieToggle.Value = False Then
        NeckCount = NeckCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_TShirtToggle_Click()
    If EX_TShirtToggle.Value = True Then
        ChestCount = ChestCount + 1
    ElseIf EX_TShirtToggle.Value = False Then
       
        ChestCount = ChestCount - 1
    End If
    UpdateCounts
End Sub
Private Sub EX_TunicToggle_Click()
    If EX_TunicToggle.Value = True Then
        ChestCount = ChestCount + 1
        HeightCount = HeightCount + 1
    ElseIf EX_TunicToggle.Value = False Then
        ChestCount = ChestCount - 1
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_DressPantsToggle_Click()
    If EX_DressPantsToggle.Value = True Then
        WaistCount = WaistCount + 1
        HipsCount = HipsCount + 1
        HeightCount = HeightCount + 1
    ElseIf EX_DressPantsToggle.Value = False Then
        WaistCount = WaistCount - 1
        HipsCount = HipsCount - 1
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_CollaredShirtToggle_Click()
    If EX_CollaredShirtToggle.Value = True Then
        NeckCount = NeckCount + 1
        ChestCount = ChestCount + 1
        HeightCount = HeightCount + 1
    ElseIf EX_CollaredShirtToggle.Value = False Then
        NeckCount = NeckCount - 1
        ChestCount = ChestCount - 1
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_WedgeToggle_Click()
    If EX_WedgeToggle.Value = True Then
        HeadCount = HeadCount + 1
    ElseIf EX_WedgeToggle.Value = False Then
        HeadCount = HeadCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_BeretToggle_Click()
    If EX_BeretToggle.Value = True Then
        HeadCount = HeadCount + 1
    ElseIf EX_BeretToggle.Value = False Then
        HeadCount = HeadCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_TillyToggle_Click()
    If EX_TillyToggle.Value = True Then
        HeadCount = HeadCount + 1
    ElseIf EX_TillyToggle.Value = False Then
        HeadCount = HeadCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_BeltToggle_Click()
    If EX_BeltToggle.Value = True Then
        WaistCount = WaistCount + 1
    ElseIf EX_BeltToggle.Value = False Then
        WaistCount = WaistCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_TiesToggle_Click()
    If EX_TiesToggle.Value = True Then
        HeightCount = HeightCount + 1
    ElseIf EX_TiesToggle.Value = False Then
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_ParkaToggle_Click()
    If EX_ParkaToggle.Value = True Then
        ChestCount = ChestCount + 1
        HipsCount = HipsCount + 1
        HeightCount = HeightCount + 1
    ElseIf EX_ParkaToggle.Value = False Then
        ChestCount = ChestCount - 1
        HipsCount = HipsCount - 1
        HeightCount = HeightCount - 1
    End If
    UpdateCounts
End Sub

Private Sub EX_InputForm_Active()

RunEveryTwoMinutes

End Sub

' # Validation

Private Function EX_DataValidation()
    ' # Validate all the input fields
    
    Dim ValidateResults(1 To 18) As Variant
    If HeadCount > 0 Then
      ValidateResults(1) = ValidateNumber(EX_HeadInput)
      ValidateResults(2) = ValidateRange(EX_HeadInput, 19, 26)
    End If
    If NeckCount > 0 Then
      ValidateResults(3) = ValidateNumber(EX_NeckInput)
      ValidateResults(4) = ValidateRange(EX_NeckInput, 12.5, 20)
    End If
    If ChestCount > 0 Then
      ValidateResults(5) = ValidateNumber(EX_ChestInput)
      ValidateResults(6) = ValidateRange(EX_ChestInput, 24, 53)
    End If
    If WaistCount > 0 Then
      ValidateResults(7) = ValidateNumber(EX_WaistInput)
      ValidateResults(8) = ValidateRange(EX_WaistInput, 30, 63)
    End If
    If HipsCount > 0 Then
      ValidateResults(9) = ValidateNumber(EX_HipsInput)
      ValidateResults(10) = ValidateRange(EX_HipsInput, 30, 68)
    End If
    If HeightCount > 0 Then
      ValidateResults(11) = ValidateNumber(EX_HeightInput)
      ValidateResults(12) = ValidateRange(EX_HeightInput, 55, 76)
    End If
    If FootLCount > 0 Then
      ValidateResults(13) = ValidateNumber(EX_FootLInput)
      ValidateResults(14) = ValidateRange(EX_FootLInput, 215, 330)
    End If
    If FootWCount > 0 Then
      ValidateResults(15) = ValidateNumber(EX_FootWInput)
      ValidateResults(16) = ValidateRange(EX_FootWInput, 85, 130)
    End If
    If HandCount > 0 Then
      ValidateResults(17) = ValidateNumber(EX_HandLInput)
      ValidateResults(18) = ValidateRange(EX_HandLInput, 6, 10)
    End If
    
    ' Check if any validation fails, early return
    Dim i As Integer
    For i = 1 To 18
        Dim ValidateResultMsg As String
        ValidateResultMsg = ValidateResults(i)
        If Not IsStringEmpty(ValidateResultMsg) Then
            MsgBox ValidateResultMsg, vbExclamation, "Input Error"
            EX_DataValidation = False
            Exit Function
        End If
    Next i
    

    EX_DataValidation = True
End Function

Sub UpdateCounts()
    ' For Chest
    If ChestCount > 0 Then
        OldExchange_Form.EX_ChestLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_ChestLabel.BackColor = &H8000000F
    End If
    
    ' For Head
    If HeadCount > 0 Then
        OldExchange_Form.EX_HeadLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_HeadLabel.BackColor = &H8000000F
    End If
    
    ' For Neck
    If NeckCount > 0 Then
        OldExchange_Form.EX_NeckLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_NeckLabel.BackColor = &H8000000F
    End If
    
    ' For Waist
    If WaistCount > 0 Then
        OldExchange_Form.EX_WaistLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_WaistLabel.BackColor = &H8000000F
    End If
    
    ' For Hips
    If HipsCount > 0 Then
        OldExchange_Form.EX_HipsLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_HipsLabel.BackColor = &H8000000F
    End If
    
    ' For FootL
    If FootLCount > 0 Then
        OldExchange_Form.EX_FootLLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_FootLLabel.BackColor = &H8000000F
    End If
    
    ' For FootW
    If FootWCount > 0 Then
        OldExchange_Form.EX_FootWLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_FootWLabel.BackColor = &H8000000F
    End If
    
    ' For Height
    If HeightCount > 0 Then
        OldExchange_Form.EX_HeightLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_HeightLabel.BackColor = &H8000000F
    End If
    
    ' For Hand
    If HandCount > 0 Then
        OldExchange_Form.EX_HandLabel.BackColor = RGB(51, 204, 204)
    Else
        OldExchange_Form.EX_HandLabel.BackColor = &H8000000F
    End If
End Sub


