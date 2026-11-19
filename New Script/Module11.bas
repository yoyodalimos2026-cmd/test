Attribute VB_Name = "Module11"
Sub Verif_Analyse_FMD_05_03_2024()

Temps_debut = Time
ActiveWorkbook.Save
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "K").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "K") = "Execution En cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "K").Interior.ColorIndex = 4



Call Verif_Recuperation_FMD_08_03_2024


ActiveWorkbook.Save
Temps_fin = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "K") = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "K").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(29, "K") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(29, "L") = Temps_fin - Temps_debut
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
End Sub
