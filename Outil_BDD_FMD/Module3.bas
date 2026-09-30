Attribute VB_Name = "Module3"
Sub Liste_Unique_FMD_29_01_2024()
Attribute Liste_Unique_FMD_29_01_2024.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Sheets("LU_FMD").Activate
    Nb_lignes_LU_FMD = ActiveWorkbook.Sheets("LU_FMD").Range("A65536").End(xlUp).Row
    If Nb_lignes_LU_FMD > 1 Then
    ActiveWorkbook.Sheets("LU_FMD").Rows(2 & ":" & ActiveWorkbook.Sheets("LU_FMD").Rows.Count).Delete
    End If
    
    ActiveWorkbook.Sheets("BDD_FMD").Activate
    ActiveWorkbook.Sheets("BDD_FMD").Columns("I:K").Select
    Selection.Copy
    Sheets("LU_FMD").Select
    ActiveWorkbook.Sheets("LU_FMD").Columns("A:A").Select
    ActiveSheet.Paste
    ActiveWorkbook.Sheets("LU_FMD").Columns("A:A").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Sheets("LU_FMD").Activate
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "C") = "Code Article ou Référence"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "D") = "CAS Number"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "E") = "Liste"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "F") = "Nom de la substance"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "G") = "Date de l'analyse"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "H") = "Date Liste"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "I") = "IMDS/Version"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "J") = "Type rapport/Compte IMDS"
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "K") = "Date d'acceptation du rapport IMDS"
    
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "C").Interior.ColorIndex = 36
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "D").Interior.ColorIndex = 36
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "E").Interior.ColorIndex = 36
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "F").Interior.ColorIndex = 36
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "G").Interior.ColorIndex = 36
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "H").Interior.ColorIndex = 36
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "I").Interior.ColorIndex = 37
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "J").Interior.ColorIndex = 37
    ActiveWorkbook.Sheets("LU_FMD").Cells(1, "K").Interior.ColorIndex = 37
    ActiveSheet.Range("$A$1:$C$10261").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    'ActiveWorkbook.Sheets("LU_FMD").Cells(1, "C") = "CAS Number"
End Sub
