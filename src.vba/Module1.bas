Sub ManualBackup()
    Dim currentDate As String
    Dim fileName As String
    Dim desktopPath As String
    Dim supplyFolderPath As String
    Dim savePath As String
    
    ' Get the current date in the "mm-dd-yyyy" format
    currentDate = Format(Now, "mm-dd-yyyy_hh_nn_ss_am/pm")
    
    ' Get the current workbook's name
    fileName = ThisWorkbook.Name
    
    ' Replace any spaces in the workbook name with underscores
    fileName = Replace(fileName, " ", "_")
    
    ' Combine the date, workbook name, and file extension
    fileName = currentDate & "Manual-" & fileName & ".xlsm" ' Add the file extension
    
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
End Sub

Function GetDesktopPath() As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Get the desktop path
    GetDesktopPath = WshShell.SpecialFolders("Desktop")
    
    Set WshShell = Nothing
End Function

Sub import()

    For ii = 2 To ThisWorkbook.Sheets("Importing").UsedRange.Rows.Count
    
        Dim nsn As String
        nsn = ActiveSheet.Cells(ii, 1).Value
        Dim addAmount As Integer
        addAmount = CInt(ActiveSheet.Cells(ii, 2).Value)

        Dim Loc As Range
        
        For Each sh In ThisWorkbook.Sheets
            If sh.Name <> "Inventory" And sh.Name <> "Importing" Then
                With sh.UsedRange
                    Set Loc = .Cells.Find(What:=nsn)
                    If Not Loc Is Nothing Then
                        Exit For
                    End If
                End With
                Set Loc = Nothing
            End If
        Next
        
        If Loc Is Nothing Then
            Exit Sub
        End If
        
        Dim Row As Integer
        Dim Col As Integer
    
        Row = Loc.Row
        Col = Loc.Column
        
        Dim QTYcol As Integer
        
        For i = Col To Col + 8
            If Loc.Worksheet.Cells(3, i).Value = "QTY" Then
                QTYcol = i
                Exit For
            End If
        Next i
        
        MsgBox Loc.Worksheet.Cells(Row, QTYcol).Value
        
        Loc.Worksheet.Cells(Row, QTYcol).Value = Loc.Worksheet.Cells(Row, QTYcol).Value + addAmount
    
    Next ii
End Sub
