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


' # Form controls
Private Sub EX_CancelButton_Click()

Unload Me

End Sub

Private Sub EX_SubmitButton_Click()
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

    Dim bPassed As Boolean
    bPassed = EX_DataValidation
    If Not bPassed Then
        Exit Sub
    End If
    
    Dim ExFormInputs(1 To 9) As MSForms.TextBox
    Set ExFormInputs(1) = EX_HeadInput
    Set ExFormInputs(2) = EX_NeckInput
    Set ExFormInputs(3) = EX_ChestInput
    Set ExFormInputs(4) = EX_WaistInput
    Set ExFormInputs(5) = EX_HipsInput
    Set ExFormInputs(6) = EX_HeightInput
    Set ExFormInputs(7) = EX_FootLInput
    Set ExFormInputs(8) = EX_FootWInput
    Set ExFormInputs(9) = EX_HandLInput

    Dim i As Integer
    For i = 1 To 9
        If Not IsStringEmpty(ExFormInputs(i).Value) Then
            ActiveSheet.Range("L" & i + 1).Value = ExFormInputs(i).Value
        End If
    Next i
    
    ReCalculateSize
    
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
      ValidateResults(6) = ValidateRange(EX_ChestInput, 24, 64)
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
        If Not IsEmpty(ValidateResults(i)) And Not ValidateResults(i) Then
            EX_DataValidation = False
            Exit Function
        End If
    Next i

    EX_DataValidation = True
End Function

Sub UpdateCounts()
    ' For Chest
    If ChestCount > 0 Then
        EXInput_Form.EX_ChestLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_ChestLabel.BackColor = &H8000000F
    End If
    
    ' For Head
    If HeadCount > 0 Then
        EXInput_Form.EX_HeadLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_HeadLabel.BackColor = &H8000000F
    End If
    
    ' For Neck
    If NeckCount > 0 Then
        EXInput_Form.EX_NeckLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_NeckLabel.BackColor = &H8000000F
    End If
    
    ' For Waist
    If WaistCount > 0 Then
        EXInput_Form.EX_WaistLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_WaistLabel.BackColor = &H8000000F
    End If
    
    ' For Hips
    If HipsCount > 0 Then
        EXInput_Form.EX_HipsLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_HipsLabel.BackColor = &H8000000F
    End If
    
    ' For FootL
    If FootLCount > 0 Then
        EXInput_Form.EX_FootLLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_FootLLabel.BackColor = &H8000000F
    End If
    
    ' For FootW
    If FootWCount > 0 Then
        EXInput_Form.EX_FootWLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_FootWLabel.BackColor = &H8000000F
    End If
    
    ' For Height
    If HeightCount > 0 Then
        EXInput_Form.EX_HeightLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_HeightLabel.BackColor = &H8000000F
    End If
    
    ' For Hand
    If HandCount > 0 Then
        EXInput_Form.EX_HandLabel.BackColor = RGB(51, 204, 204)
    Else
        EXInput_Form.EX_HandLabel.BackColor = &H8000000F
    End If
End Sub

Private Sub UserForm_Click()

End Sub