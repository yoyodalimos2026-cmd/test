Attribute VB_Name = "Module1"
Dim cell As Range
Dim fileName As String
Dim fullfilename As String
Dim rownumber As Long
Dim columnnumber As Long
Dim wb As Workbook
Dim wsRange As Worksheet
Dim wsHome As Worksheet
Public nbTraite As Integer
'***************************************'
Sub CreateHyperLinks()
'***************************************'
    Dim fileobj As Object
    Dim folderobj As Object
    Dim fpath As String
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
    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    '***************************************'
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    '***************************************'
    If fDialog.Show = -1 Then
        fpath = fDialog.SelectedItems(1)
    End If
    '***************************************'
    Set fDialog = Nothing
    '***************************************'
    If fpath = "" Then Exit Sub
    '***************************************'
    '***************************************'
    filetype = InputBox("Enter file type")
    '***************************************'
    For Each cell In rng
    '***************************************'
        fileName = cell.Value
        '***************************************'
        If Right(fileName, Len(filetype)) = filetype And cell.Hyperlinks.Count = 0 Then
        '***************************************'
            fullfilename = fileName
            rownumber = cell.Row
            columnnumber = cell.Column
            '***************************************'
            fileFound = Recurse(fpath)
            '***************************************'
            If Not fileFound Then
                cell.Interior.Color = vbRed
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

Function Recurse(sPath As String) As Boolean
'***************************************'
    Dim FSO As New FileSystemObject
    Dim myFolder As folder
    Dim mySubFolder As folder
    Dim myFile As file
    Dim hyper As String
    '***************************************'
    Set myFolder = FSO.GetFolder(sPath)
    '***************************************'
    For Each myFile In myFolder.Files
    '***************************************'
        If StrComp(myFile.Name, fullfilename, vbTextCompare) = 0 Then
        '***************************************'
            nbTraite = nbTraite + 1
            hyper = "./" & myFolder.Name & "/" & fullfilename
            Cells(rownumber, columnnumber).Formula = "=HYPERLINK(""" & hyper & """, """ & fullfilename & """)"
            '***************************************'
            Recurse = True
            Exit Function
            '***************************************'
        End If
        '***************************************'
    Next
    '***************************************'
    For Each mySubFolder In myFolder.SubFolders
    '***************************************'
        If Recurse(mySubFolder.Path) Then
            Recurse = True
            Exit Function
        End If
        '***************************************'
    Next
    '***************************************'
    Recurse = False
    '***************************************'
End Function

