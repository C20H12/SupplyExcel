' comment

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim currentDate As String
    Dim fileName As String
    Dim desktopPath As String
    Dim supplyFolderPath As String
    Dim savePath As String
    
    ' Get the current date in the "mm-dd-yyyy" format
    currentDate = Format(Date, "mm-dd-yyyy")
    
    ' Get the current workbook's name
    fileName = ThisWorkbook.Name
    
    ' Replace any spaces in the workbook name with underscores
    fileName = Replace(fileName, " ", "_")
    
    ' Combine the date, workbook name, and file extension
    fileName = currentDate & "-" & fileName & ".xlsm" ' Add the file extension
    
    ' Get the user's desktop path
    desktopPath = GetDesktopPath
    
    supplyFolderPath = desktopPath & "\Supply 2.0\"
    strsupplyFolderExists = Dir(supplyFolderPath)

   If strsupplyFolderExists = "" Then
        MkDir supplyFolderPath
    End If
    
    ' Define the full save path
    savePath = supplyFolderPath & fileName
    
    ' Save a copy of the workbook with the constructed file name
    ThisWorkbook.SaveCopyAs savePath
    
    ' Close the newly saved copy without saving changes to the original workbook
    'Workbooks.Open savePath
    'ActiveWorkbook.Close SaveChanges:=False
End Sub

Function GetDesktopPath() As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Get the desktop path
    GetDesktopPath = WshShell.SpecialFolders("Desktop")
    
    Set WshShell = Nothing
End Function