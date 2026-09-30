Attribute VB_Name = "Module1"
Sub CopyFiles()
    Dim folderPath As String
    Dim targetFolder As String
    Dim fileNames As Range
    Dim fileName As String
    Dim file As Variant
    Dim oFSO As Object
    Dim oDIR As Object
    Dim ws1 As Worksheet
    Dim dbt As Date, fin As Date, temps As Date
    
    dbt = Time
    
    Set ws1 = ActiveWorkbook.Sheets("Home")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Set folder path and target folder
    folderPath = InputBox("Enter the folder path to search for files:", "Search Folder Path")
    targetFolder = InputBox("Enter the folder path to copy files in:", "Target Path")
    
    dbt = Time
    ws1.Cells(2, 8) = folderPath
    ws1.Cells(3, 8) = targetFolder
    
    Set oDIR = oFSO.GetFolder(targetFolder)
    Z = oDIR.Files.Count
    
    ' Set file names range
    Set fileNames = Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
    
    x = 1
    y = 0
    ' Loop through file names
    For Each file In fileNames
        
        x = x + 1
        
        fileName = file.Value
        
        ' Check if file exists in folder
        If Dir(folderPath & "\" & fileName) <> "" Then
            ' Copy file to target folder
            FileCopy folderPath & "\" & fileName, targetFolder & "\" & fileName
            ws1.Cells(x, 1).Interior.ColorIndex = 4
            
        Else
            y = y + 1
            ws1.Cells(x, 1).Interior.ColorIndex = 3
        End If
    Next file
    
    ws1.Cells(5, 8) = oDIR.Files.Count - Z
    ws1.Cells(6, 8) = y
    
    fin = Time
    temps = dbt - fin
    ws1.Cells(8, 8) = Format(temps, "hh:mm:ss")
    
End Sub
