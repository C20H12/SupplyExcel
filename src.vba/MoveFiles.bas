Sub CopyWorksheetsFromExternalWorkbook()
    Dim sourceFilePath As String
    Dim sourceWorkbook As Workbook
    Dim currentWorkbook As Workbook
    Dim ws As Worksheet
    
    ' Set the path to the source workbook
    sourceFilePath = "C:\Users\david\OneDrive\Desktop\Supply_2.0_v_0.1.1.0.xlsm"
    
    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    
    ' Set the current workbook (the one where you want to insert the sheets)
    Set currentWorkbook = ThisWorkbook
    
    ' Loop through each worksheet in the source workbook
    For Each ws In sourceWorkbook.Worksheets
        ' Check if the sheet name is not one of the excluded names
        If ws.Name <> "Importing" And ws.Name <> "Menu" And ws.Name <> "Template" Then
            ' Copy the sheet to the current workbook
            ws.Copy After:=currentWorkbook.Sheets(currentWorkbook.Sheets.count)
        End If
    Next ws
    
    ' Close the source workbook without saving changes
    sourceWorkbook.Close SaveChanges:=False
End Sub