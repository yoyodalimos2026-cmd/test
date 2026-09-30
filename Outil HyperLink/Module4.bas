Attribute VB_Name = "Module4"
Dim cell As Range
Dim fileName As String
Dim fullfilename As String
Dim rownumber As Long
Dim columnnumber As Long
Dim wb As Workbook
Dim wsRange As Worksheet
Dim wsHome As Worksheet
Public nbTraite As Integer
Public fileCollection As Collection
'***************************************'
Sub CreateHyperLinks()
'***************************************'
    Dim fileobj As Object
    Dim folderobj As Object
    Dim fpath As String
    Dim hyper As String
    Dim filep As Variant
    Dim fDialog As FileDialog
    Dim filetype As String
    Dim rng As Range
    Dim fileFound As Boolean
    '***************************************'
    Set wb = ThisWorkbook
    Set wsRange = wb.Sheets("Range")
    Set wsHome = wb.Sheets("Home")
    '***************************************'
    launchTime = Time()
    wsHome.Cells(4, "A") = "En Cours d'éxecution"
    wsHome.Cells(4, "A").Interior.ColorIndex = 4
    '***************************************'
    wsRange.Activate
    Application.Calculation = xlCalculationManual
    nbTraite = 0
    '***************************************'
    On Error Resume Next
    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    If rng Is Nothing Then Exit Sub
    '***************************************'
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    '***************************************'
    If fDialog.Show = -1 Then
        fpath = fDialog.SelectedItems(1)
    End If
    '***************************************'
    'If fpath = "" Then Exit Sub
    Set fDialog = Nothing
    '***************************************'
    If fpath = "" Then Exit Sub
    '***************************************'
    If fileCollection Is Nothing Then
        Set fileCollection = New Collection
    End If
    Call ListFilesInDirectory(fileCollection, fpath)
    '***************************************'
    filetype = InputBox("Enter file type")
    '***************************************'
    For Each cell In rng
    '***************************************'
        If Not IsEmpty(cell) Then
    '***************************************'
            fileName = cell.Value
            '***************************************'
            If Right(fileName, Len(filetype)) = filetype And cell.Hyperlinks.Count = 0 Then
            '***************************************'
                fullfilename = fileName
                rownumber = cell.Row
                columnnumber = cell.Column
                '***************************************'
                fileFound = False
                For Each filep In fileCollection
                    If InStr(filep, fileName) > 0 Then
                        'hyper = filep
                        hyper = "." & ExtractSubstring(CStr(filep))
                        Debug.Print hyper
                        wsRange.Cells(rownumber, columnnumber).Formula = "=HYPERLINK(""" & hyper & """, """ & fileName & """)"
                        'cell.Hyperlinks.Add Anchor:=cell, Address:=hyper, TextToDisplay:=fileName
                        fileFound = True
                        nbTraite = nbTraite + 1
                        Exit For
                    'Else
                        'wsRange.Cells(rownumber, columnnumber).Interior.Color = vbRed
                    End If
                Next filep
                '***************************************'
                If Not fileFound Then
                    cell.Interior.Color = vbRed
                End If
                '***************************************'
            End If
            '***************************************'
        End If
        '***************************************'
    Next cell
    '***************************************'
    Application.Calculation = xlCalculationAutomatic
    '***************************************'
    wsHome.Cells(4, "A") = "Execution Terminé"
    wsHome.Cells(4, "A").Interior.ColorIndex = 9
    endTime = Time()
    passedTime = Format((endTime - launchTime), "hh:mm:ss")
    wsHome.Cells(5, "B").Value = passedTime
    wsHome.Cells(5, "A").Value = Date & " à " & Format(endTime, "hh:mm:ss")
    wsHome.Cells(7, "B").Value = nbTraite
    wb.Save
    '***************************************'
End Sub

Sub ListFilesInDirectory(fileCollection As Collection, folderPath As String)
    Dim fileSystem As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    Set folder = fileSystem.GetFolder(folderPath)

    For Each file In folder.Files
        fileCollection.Add file.Path
    Next file

    For Each subFolder In folder.SubFolders
        ListFilesInDirectory fileCollection, subFolder.Path
    Next subFolder

    Set fileSystem = Nothing
End Sub
Function ExtractSubstring(inputString As String) As String
    Dim pos As Integer
    Dim secondLastPos As Integer
    pos = InStrRev(inputString, "\")
    secondLastPos = InStrRev(inputString, "\", pos - 1)
    ExtractSubstring = Mid(inputString, secondLastPos)
    
End Function
