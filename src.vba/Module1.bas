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
    ' stores size, name, old_amount, new amount
    Dim changesCount As Integer
    changesCount = 0
    Dim changes() As Variant
    Dim qtyCells() As Variant
    Dim newqty() As Variant

    For ii = 2 To ThisWorkbook.Sheets("Importing").UsedRange.Rows.Count
        
        changesCount = changesCount + 1
        ReDim Preserve changes(changesCount)
        ReDim Preserve qtyCells(changesCount)
        ReDim Preserve newqty(changesCount)
    
        Dim nsn As String
        nsn = ActiveSheet.Cells(ii, 1).Value
        
        ' if misinputted a space the used rows are more than it should be
        If Len(Trim(nsn)) = 0 Then
            Exit For
        End If
        
        Dim addAmount As Integer
        addAmount = CInt(ActiveSheet.Cells(ii, 2).Value)

        Dim Loc As Range
        
        For Each sh In ThisWorkbook.Sheets
            If sh.Name <> "Inventory" And sh.Name <> "Importing" Then
                With sh.UsedRange
                    Set Loc = .Cells.Find(What:=nsn)
                    ' is found
                    If Not Loc Is Nothing Then
                        Exit For
                    End If
                End With
                Set Loc = Nothing
            End If
        Next
        
        If Loc Is Nothing Then
            changes(changesCount) = "Invalid NSN"
            GoTo continue
        Else
            changes(changesCount) = Loc.Offset(0, 1).Value & vbTab & Loc.Worksheet.Cells(1, Loc.Column) & vbTab
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
        
       ' MsgBox Loc.Worksheet.Cells(Row, QTYcol).Value
        
        Dim oldamount As Integer
        oldamount = Loc.Worksheet.Cells(Row, QTYcol).Value
        changes(changesCount) = changes(changesCount) & "From " & oldamount & " to " & oldamount + addAmount
        Set qtyCells(changesCount) = Loc.Worksheet.Cells(Row, QTYcol)
        newqty(changesCount) = oldamount + addAmount
        'Loc.Worksheet.Cells(Row, QTYcol).Value = Loc.Worksheet.Cells(Row, QTYcol).Value + addAmount
        
continue:
    Next ii
    
    Dim dispStr As String
    dispStr = "These Will Be Modified: "
    
    For i = 1 To changesCount
        dispStr = dispStr & vbNewLine & changes(i)
    Next i
    
    If MsgBox(dispStr, vbYesNo) = vbYes Then
        ActiveSheet.Range("A2:B" & ThisWorkbook.Sheets("Importing").UsedRange.Rows.Count).Delete
        On Error Resume Next
        For i = 1 To changesCount
            qtyCells(i).Value = newqty(i)
        Next i
    End If
End Sub
