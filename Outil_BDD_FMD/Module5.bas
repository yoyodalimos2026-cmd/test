Attribute VB_Name = "Module5"
Sub Analyse_FMD()

Temps_debut = Time
ActiveWorkbook.Save
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "D").Interior.ColorIndex = 2
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "D") = "Execution En cours"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "D").Interior.ColorIndex = 4

ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(30, "B") = (ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(32, "B") - ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(31, "B")) + 1
Call Conversion_xlsx_29_01_2024
Call Recuperation_FMD_29_01_2024
Call MAJ_Etat_FMD_BDD


Temps_fin = Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "D") = "Execution Terminée"
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(28, "D").Interior.ColorIndex = 9
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(29, "D") = Date & " a " & Time
ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(29, "E") = Temps_fin - Temps_debut
ActiveWorkbook.Sheets("Macro_MAJ_Document").Activate
ActiveWorkbook.Save
End Sub
