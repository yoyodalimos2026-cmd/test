Attribute VB_Name = "Module15"
Sub Print_pdf()

Dim FichierAImprimer As Variant

Application.ActivePrinter = "PDFCreator sur Ne00:"
'SetDefaultPrinter "PDFCreator"
Current_path = ThisWorkbook.Path
Dossier_FMD = ActiveWorkbook.Sheets("Macro_MAJ_Document").Cells(7, "B")
'Application.DefaultFilePath = Current_path & "\" & Dossier_FMD & "\" & "Vishay_FMD_293D475X0025C2TE3_1.pdf"

'Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
'Application.FileDialog(msoFileDialogOpen).InitialFileName = Current_path
'intChoice = Application.FileDialog(msoFileDialogOpen).Show

'Variable_Path = Application.DefaultFilePath

For u = 56 To 56

FichierAImprimer = Current_path & "\" & Dossier_FMD & "\" & ActiveWorkbook.Sheets("LU_FMD").Cells(u, "A") '<-- chemin complet du fichier à imprimer
Debug.Print FichierAImprimer
CreateObject("Shell.Application").Namespace(0).ParseName(FichierAImprimer).InvokeVerb ("Print")
'Application.DefaultFilePath = Current_path & "\" & Dossier_FMD & "\" & "Vishay_FMD_293D475X0025C2TE3_1.pdf"
Next u

End Sub
