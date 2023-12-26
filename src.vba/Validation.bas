Function ValidateText(InputBox As MSForms.TextBox) As Boolean
    If Not IsLettersOrUnderscores(InputBox.Value) Then
        MsgBox "Please enter letters or letters with spaces for " & InputBox.Tag, vbExclamation, "Input Error"
        InputBox.SetFocus
        ValidateText = False
    Else
        ValidateText = True
    End If
End Function

Function ValidateNumber(InputBox As MSForms.TextBox) As Boolean
    If Not IsNumeric(InputBox.Value) Then
        MsgBox "Please enter a numeric value for " & InputBox.Tag, vbExclamation, "Input Error"
        InputBox.SetFocus
        ValidateNumber = False
    Else
        ValidateNumber = True
    End If
End Function

Function ValidateDate(InputBox As MSForms.TextBox) As Boolean
    If Not IsDate(InputBox.Value) Then
        MsgBox "Please enter a date(mm/dd/yyyy) For " & InputBox.Tag, vbExclamation, "Input Error"
        InputBox.SetFocus
        ValidateDate = False
    Else
        ValidateDate = True
End Function

Function ValidateRange(InputBox As MSForms.TextBox, Min As Double, Max As Double) As Boolean
    If IsStringEmpty(InputBox.Value) Or Not IsNumeric(InputBox.Value) Then
        ValidateRange = False
        Exit Function
    End If
    Dim InputAsNum As Double
    InputAsNum = CDbl(InputBox.Value)
    If InputAsNum < Min Or InputAsNum > Max Then
        MsgBox "Please enter a number between " & Min & " and " & Max & " for " & InputBox.Tag & " Measurement.", vbExclamation, "Input Error"
        InputBox.SetFocus
        ValidateRange = False
    Else
        ValidateRange = True
    End If
End Function

Function ValidateCustom(InputBox As MSForms.TextBox, Condition As Boolean, Message As String) As Boolean
    If Condition Then
        MsgBox Message, vbExclamation, "Input Error"
        InputBox.SetFocus
        ValidateCustom = False
    Else
        ValidateCustom = True
    End If
End Function