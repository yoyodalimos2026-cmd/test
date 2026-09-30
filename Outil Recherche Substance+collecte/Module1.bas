Attribute VB_Name = "Module1"
Sub Collecte_BDD()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsHome As Worksheet
    Dim wsOut As Worksheet
    Dim outRow As Long
    Dim filesProcessed As Long
    
    Dim startTime As Double
    Dim endTime As Double
    Dim duration As Double
    
    ' Set references to the sheets in the current macro workbook (.xlsm)
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets("Home")
    Set wsOut = ThisWorkbook.Sheets("BDD")
    On Error GoTo 0
    
    If wsHome Is Nothing Or wsOut Is Nothing Then
        MsgBox "Error: Could not find sheets named exactly 'Home' and 'BDD' in this workbook.", vbCritical, "Missing Sheets"
        Exit Sub
    End If
    
    ' Record start time
    startTime = Now
    
    ' Prompt user to select the target folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Folder containing the Excel files"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "Folder selection cancelled.", vbExclamation, "Cancelled"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    
    ' Ensure folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Setup the BDD Sheet
    wsOut.Cells.Clear ' Clear previous data
    
    ' Add Filename Header
    wsOut.Range("A1").Value = "Filename"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Interior.Color = 5296274
    outRow = 2
    filesProcessed = 0
    
    ' Dictionary to keep track of unique List Names and their assigned column indices
    Dim dictListCols As Object
    Set dictListCols = CreateObject("Scripting.Dictionary")
    dictListCols.CompareMode = vbTextCompare ' Case-insensitive matching for List Names
    Dim nextCol As Long
    nextCol = 2 ' Start creating list columns from Column B
    
    ' Optimize macro performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through all Excel files in the folder
    fileName = Dir(folderPath & "*.xls*")
    
    Do While fileName <> ""
        ' Prevent the macro from trying to open itself if it is in the same folder
        If fileName <> ThisWorkbook.Name Then
            
            ' Attempt to open the file in Read-Only mode
            On Error Resume Next
            Set wbSource = Workbooks.Open(folderPath & fileName, ReadOnly:=True, UpdateLinks:=False)
            On Error GoTo 0
            
            If Not wbSource Is Nothing Then
                ' Data is located on the first sheet of the workbook
                Set wsSource = wbSource.Sheets(1)
                
                ' Find the column indices dynamically based on the header names
                Dim colCAS As Long, colSub As Long, colList As Long
                colCAS = 0: colSub = 0: colList = 0
                
                Dim lastCol As Long
                lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
                
                Dim c As Long
                For c = 1 To lastCol
                    Select Case LCase(Trim(wsSource.Cells(1, c).Value))
                        Case "cas number": colCAS = c
                        Case "substance name": colSub = c
                        Case "list name": colList = c
                    End Select
                Next c
                
                ' If CAS and List columns are found, process the data
                If colCAS > 0 And colList > 0 Then
                    ' Dictionary for the current file's List Names and their concatenated CAS numbers
                    Dim dictFileLists As Object
                    Set dictFileLists = CreateObject("Scripting.Dictionary")
                    dictFileLists.CompareMode = vbTextCompare
                    
                    Dim lastRowCAS As Long, lastRowList As Long
                    lastRowCAS = wsSource.Cells(wsSource.Rows.Count, colCAS).End(xlUp).Row
                    lastRowList = wsSource.Cells(wsSource.Rows.Count, colList).End(xlUp).Row
                    
                    Dim maxRow As Long
                    maxRow = Application.WorksheetFunction.Max(lastRowCAS, lastRowList)
                    
                    Dim r As Long
                    Dim valCAS As String, valList As String
                    
                    ' Loop through each row to group CAS numbers by List Name
                    For r = 2 To maxRow
                        valCAS = Trim(CStr(wsSource.Cells(r, colCAS).Value))
                        valList = Trim(CStr(wsSource.Cells(r, colList).Value))
                        
                        If valList <> "" Then
                            If dictFileLists.Exists(valList) Then
                                If valCAS <> "" Then
                                    If dictFileLists(valList) <> "" Then
                                        dictFileLists(valList) = dictFileLists(valList) & "//" & valCAS
                                    Else
                                        dictFileLists(valList) = valCAS
                                    End If
                                End If
                            Else
                                dictFileLists(valList) = valCAS
                            End If
                        End If
                    Next r
                    
                    ' Write the extracted data to the BDD sheet
                    wsOut.Cells(outRow, 1).Value = fileName
                    
                    ' Loop through the unique lists found in this specific file
                    Dim k As Variant
                    For Each k In dictFileLists.Keys
                        Dim listName As String
                        listName = CStr(k)
                        
                        ' If this List Name hasn't been seen before, create a new column header
                        If Not dictListCols.Exists(listName) Then
                            dictListCols.Add listName, nextCol
                            wsOut.Cells(1, nextCol).Value = listName
                            wsOut.Cells(1, nextCol).Font.Bold = True
                            wsOut.Cells(1, nextCol).Interior.Color = 5296274
                            nextCol = nextCol + 1
                        End If
                        
                        ' Write the concatenated CAS numbers for this list into the correct column
                        wsOut.Cells(outRow, dictListCols(listName)).Value = dictFileLists(listName)
                    Next k
                    
                    outRow = outRow + 1
                    filesProcessed = filesProcessed + 1
                    
                Else
                    ' Log an error if the headers are missing
                    wsOut.Cells(outRow, 1).Value = fileName
                    wsOut.Cells(outRow, 2).Value = "ERROR: Required headers (CAS Number or List Name) not found on Sheet 1"
                    outRow = outRow + 1
                    filesProcessed = filesProcessed + 1
                End If
                
                ' Close source file without saving
                wbSource.Close SaveChanges:=False
                Set wbSource = Nothing
            Else
                ' Log an error if the file is corrupted or password protected
                wsOut.Cells(outRow, 1).Value = fileName
                wsOut.Cells(outRow, 2).Value = "ERROR: Could not open file"
                outRow = outRow + 1
                filesProcessed = filesProcessed + 1
            End If
        End If
        
        ' Get the next file
        fileName = Dir()
    Loop
    
    ' Finalize and format the BDD Sheet
    wsOut.UsedRange.Columns.AutoFit
    
    ' Re-enable Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Record end time and calculate duration
    endTime = Now
    duration = endTime - startTime
    
    ' Convert to total seconds to safely calculate hours without rolling over at 24h
    Dim totalSecs As Long
    totalSecs = duration * 86400 ' 86400 seconds in a day
    Dim h As Long, m As Long, s As Long
    h = totalSecs \ 3600
    m = (totalSecs Mod 3600) \ 60
    s = totalSecs Mod 60
    Dim durationStr As String
    durationStr = h & "h " & m & "m " & s & "s"
    
    ' Update the "Home" sheet with execution statistics
    wsHome.Range("C3").Value = "Last Execution Time:"
    wsHome.Range("D3").Value = Format(startTime, "yyyy-mm-dd hh:mm:ss")
    
    wsHome.Range("C4").Value = "Execution Duration:"
    wsHome.Range("D4").Value = durationStr
    
    wsHome.Range("C5").Value = "Total Files Processed:"
    wsHome.Range("D5").Value = filesProcessed
    
    wsHome.Range("C3:D5").Font.Bold = True
    wsHome.Columns("C:D").AutoFit
    
    MsgBox "Processing complete! Check the 'BDD' sheet for data.", vbInformation, "Success"
End Sub

