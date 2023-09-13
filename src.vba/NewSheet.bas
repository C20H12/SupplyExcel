Sub CreateNewCadetSheet(ByVal sNewSheetName As String)
'
' Test_1 Macro
'
' This is for creating a new Sheets File'

'
    Sheets("Template").Select
    Range("A1:L36").Select
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
    Columns("E:E").ColumnWidth = 15
    Range("J5").Select
   
End Sub
