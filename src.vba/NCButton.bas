Sub NCInput_Button()

    If Not NCInput_Form.Visible Then
        NCInput_Form.Show

    End If

End Sub
Sub Range_Test()
'
' Range_Test Macro
'

'
    Range("L2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("L2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:="19", Formula2:="26"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub