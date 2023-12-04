Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    ' check the cells from e6 to e24 (sizes), if changed, run code
    Dim RangeToCheck As Range
    Set RangeToCheck = sh.Range("E6:E24")

    If Not Application.Intersect(RangeToCheck, Target) Is Nothing Then
        Dim NSNResult As String
        NSNResult = GetNSNFromSize(Target.Offset(0, -3).Value, Target.Value, sh.Range("G4").Value = "Male")
        If IsStringEmpty(NSNResult) Then
            NSNResult = "Invalid size"
        End If
        ' move left 3 col to A, then insert the result
        Target.Offset(0, -4).Value = NSNResult
    End If
End Sub


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