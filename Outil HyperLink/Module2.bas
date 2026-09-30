Attribute VB_Name = "Module2"
Sub ListFilesInDirectory(fileCollection As Collection, folderPath As String)
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object

    ' Initialize the FileSystemObject
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    ' Get the folder object
    Set folder = fileSystem.GetFolder(folderPath)

    ' Loop through each file in the folder
    For Each file In folder.Files
        ' Add the relative file path to the collection
        fileCollection.Add file.Path
    Next file

    ' Clean up
    Set fileSystem = Nothing
End Sub

Sub CreateHyperlinksFromCollection()
    Dim fileCollection As Collection
    Dim filePath As Variant
    Dim fileName As String
    Dim cell As Range
    Dim folderPath As String

    ' Prompt user to enter the directory path
    folderPath = InputBox("Enter the directory path:")

    ' Initialize the collection
    Set fileCollection = New Collection

    ' Populate the file collection
    ListFilesInDirectory fileCollection, folderPath

    ' Loop through each cell in column A
    For Each cell In ThisWorkbook.Sheets(2).Range("A1:A" & ThisWorkbook.Sheets(2).Cells(Rows.Count, 1).End(xlUp).Row)
        fileName = cell.Value
        ' Loop through each file path in the collection
        For Each filePath In fileCollection
            ' Check if the file name matches
            If InStrRev(filePath, "\") > 0 Then
                If fileName = Mid(filePath, InStrRev(filePath, "\") + 1) Then
                    ' Create a relative hyperlink
                    ThisWorkbook.Sheets(1).Hyperlinks.Add Anchor:=cell, Address:=filePath, TextToDisplay:=fileName
                    Exit For
                End If
            End If
        Next filePath
    Next cell

    ' Clean up
    Set fileCollection = Nothing
End Sub

