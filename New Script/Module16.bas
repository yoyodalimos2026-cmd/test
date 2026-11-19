Attribute VB_Name = "Module16"
Sub PrintWithPDFConverter()
    Dim filePath As String
    Dim acrobatPath As String
    Dim acrobatPrintCommand As String
    
    Nb_Lignes_Data = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
    Current_path = ThisWorkbook.Path
    Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B")
    ' Specify the file path of the file you want to print
    Ligne_debut = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "B")
    Ligne_Fin = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(32, "B")
    
    For t = Ligne_debut To Ligne_Fin
    
        If InStr(ActiveWorkbook.Sheets("LU_FMD").Cells(t, "B"), "FMD Format IPC") <> 0 Then
    
        filePath = Current_path & "\" & Dossier_FMD & "\" & ActiveWorkbook.Sheets("LU_FMD").Cells(t, "A")
        'Debug.Print filePath
        ' Specify the path to Adobe Acrobat Reader or Acrobat Pro
        acrobatPath = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

        ' Specify the print command for Adobe Acrobat
        acrobatPrintCommand = "/t """ & filePath & """"

        ' Execute the print command with Adobe Acrobat
        '?Application.ActivePrinter <== dans la fenetre d'excution
        Application.ActivePrinter = "PDFCreator sur Ne00:"
        Shell acrobatPath & " " & acrobatPrintCommand, vbNormalFocus
        Application.Wait Now + TimeValue("0:00:10")
    
        ' Show a message box indicating that printing has started
        'MsgBox "Printing with PDF converter has started."
        End If
    Next t
End Sub

