Function IsLettersOrUnderscores(ByVal inputStr As String) As Boolean
    Dim i As Integer
    Dim c As String
    
    IsLettersOrUnderscores = True
    
    For i = 1 To Len(inputStr)
        c = Mid(inputStr, i, 1)
        If Not (c Like "[A-Za-z_ ]") Then
            IsLettersOrUnderscores = False
            Exit Function
        End If
    Next i
End Function

Function IsNumbersOrLetters(ByVal inputStr As String) As Boolean
    Dim i As Integer
    Dim c As String
    
    IsNumbersOrLetters = True
    
    For i = 1 To Len(inputStr)
        c = Mid(inputStr, i, 1)
        If Not (c Like "[A-Za-z0-9]") Then
            IsNumbersOrLetters = False
            Exit Function
        End If
    Next i
End Function

Function IsStringEmpty(Inp As String)
    IsStringEmpty = (Len(Trim(Inp)) = 0)
End Function

Function GetUUID() As String
    ' Generates an 8 digit random ID
    Dim uuid(7) As Integer
    Randomize
    For i = 0 To 7
        uuid(i) = Int(Rnd() * 16)
    Next i
    uuid(6) = uuid(6) And (Not 4)
    uuid(6) = uuid(6) Or 8
    Dim uuidString As String
    For i = 0 To 7
        uuidString = uuidString & Hex(uuid(i))
    Next i
    GetUUID = uuidString
End Function

Sub QuickSort(arr() As Integer, left As Integer, right As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim pivot As String
    Dim temp As String
    
    i = left
    j = right
    pivot = arr((left + right) \ 2)
    
    While i <= j
        While arr(i) < pivot And i < right
            i = i + 1
        Wend
        
        While pivot < arr(j) And j > left
            j = j - 1
        Wend
        
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend
    
    If left < j Then
        QuickSort arr, left, j
    End If
    
    If i < right Then
        QuickSort arr, i, right
    End If
End Sub

Function ArrayToString(ByRef arr() As Variant) As String
    Dim out As String
    out = ""
    For i = 0 To UBound(arr)
        out = out & arr(i)
        MsgBox arr(i)
    Next i
    ArrayToString = out
End Function

Function isSpecialSheet(SheetName As String)

    isSpecialSheet = (SheetName = "Menu" Or SheetName = "Importing" Or SheetName = "Pickup" Or SheetName = "Template" Or SheetName = "Master" Or SheetName = "Import Sheets")

End Function