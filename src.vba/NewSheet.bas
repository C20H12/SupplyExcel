Sub CreateNewCadetSheet(ByVal sNewSheetName As String)
'
' Test_1 Macro
'
' This is for creating a new Sheets File'

'
    Sheets("Template").Select
    Range("A1:L26").Select
    Selection.Copy
    Sheets.Add.Name = sNewSheetName
    Sheets(sNewSheetName).Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("P24").Select
    Columns("K:K").ColumnWidth = 11.43
    Columns("A:A").ColumnWidth = 17.86
    Range("O14").Select
    Columns("L:L").ColumnWidth = 12.43
    
    Columns("H:H").ColumnWidth = 1.78
    Range("H6").Select
    Application.Goto Reference:="R6C7"
    Range("I6").Select
    Columns("G:G").ColumnWidth = 16
    Range("J5").Select
    
    ' # sos button
    Dim sosButton As Button
    Dim BtnLeft As Double, BtnTop As Double
    Dim BtnWidth As Double, BtnHeight As Double
    
    ' Set button dimensions and position
    BtnLeft = Range("K25").left
    BtnTop = Range("K25").Top
    BtnWidth = Range("L25").left + Range("L25").Width - Range("K25").left
    BtnHeight = Range("K26").Top + Range("K26").Height - Range("K25").Top
    
    ' Add the button
    Set sosButton = ActiveSheet.Buttons.Add(BtnLeft, BtnTop, BtnWidth, BtnHeight)
    
    ' Set button properties
    With sosButton
        .Caption = "S.O.S."
        .OnAction = "terminate"
    End With
    
    ' # resize button
    Dim ResizeButton As Button
    
    ' Set button dimensions and position
    BtnLeft = Range("K12").left
    BtnTop = Range("K12").Top
    BtnWidth = Range("L12").left + Range("L12").Width - Range("K12").left
    BtnHeight = Range("K13").Top + Range("K13").Height - Range("K12").Top
    
    ' Add the button
    Set ResizeButton = ActiveSheet.Buttons.Add(BtnLeft, BtnTop, BtnWidth, BtnHeight)
    
    ' Set button properties
    With ResizeButton
        .Caption = "Resize"
        .OnAction = "ReCalculateSize"
    End With

    ' # exchange button
    Dim ExchangeButton As Button
    
    ' Set button dimensions and position
    BtnLeft = Range("K17").left
    BtnTop = Range("K17").Top
    BtnWidth = Range("L17").left + Range("L17").Width - Range("K17").left
    BtnHeight = Range("K18").Top + Range("K18").Height - Range("K17").Top
    
    ' Add the button
    Set ExchangeButton = ActiveSheet.Buttons.Add(BtnLeft, BtnTop, BtnWidth, BtnHeight)
    
    ' Set button properties
    With ExchangeButton
        .Caption = "Exchange Item"
        .OnAction = "ExchangeButton"
    End With
End Sub
