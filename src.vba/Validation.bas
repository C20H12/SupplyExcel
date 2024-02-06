Function ValidateText(InputBox As MSForms.TextBox) As String
    If IsStringEmpty(InputBox.Value) Or Not IsLettersOrUnderscores(InputBox.Value) Then
        ValidateText = "Please enter letters or letters with spaces for " & InputBox.Tag
    Else
        ValidateText = ""
    End If
End Function

Function ValidateNumber(InputBox As MSForms.TextBox) As String
    If Not IsNumeric(InputBox.Value) Then
        ValidateNumber = "Please enter a numeric value for " & InputBox.Tag
    Else
        ValidateNumber = ""
    End If
End Function

Function ValidateRange(InputBox As MSForms.TextBox, Min As Double, Max As Double) As String
    If IsStringEmpty(InputBox.Value) Or Not IsNumeric(InputBox.Value) Then
        ValidateRange = "Please enter a numeric value for " & InputBox.Tag
        Exit Function
    End If
    Dim InputAsNum As Double
    InputAsNum = CDbl(InputBox.Value)
    If InputAsNum < Min Or InputAsNum > Max Then
        ValidateRange = "Please enter a number between " & Min & " and " & Max & " for " & InputBox.Tag & " Measurement."
    Else
        ValidateRange = ""
    End If
End Function

Function ValidateCustom(InputBox As MSForms.TextBox, Condition As Boolean, Message As String) As String
    If Condition Then
        ValidateCustom = Message
    Else
        ValidateCustom = ""
    End If
End Function


Function ValidateBlank(InputBox As MSForms.TextBox) As String
    If IsStringEmpty(InputBox.Value) Then
        MsgBox "Please enter a value for " & InputBox.Tag, vbExclamation, "Input Error"
    Else
        ValidateBlank = ""
    End If
End Function
