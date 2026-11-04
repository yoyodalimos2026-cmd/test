Attribute VB_Name = "Module7"

Option Explicit

' Ne pas oublier de cocher la bibliothèqe PDFCreator
' dans le menu Outils / Références de Visual Basic Editor

' API Windows pour faire une temporisation en millisecondes
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


' Constantes pour les temporisations
Private Const maxTime = 10    ' en secondes
Private Const sleepTime = 250 ' en millisecondes
Const deblin As Integer = 5
 
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

 
Function FileExists(ByVal sFileName As String) As Boolean

    Dim FSO

    Set FSO = CreateObject("Scripting.FileSystemObject")

    FileExists = FSO.FileExists(sFileName)

End Function

 
Function WorksheetExists(ByVal Name As String, Optional wb As Workbook) As Boolean

    Dim ws As Worksheet

 
    WorksheetExists = False

    If wb Is Nothing Then Set wb = ActiveWorkbook

 
    For Each ws In wb.Worksheets

        If ws.Name = Name Then

            WorksheetExists = True

            Exit For

        End If

    Next

End Function
Sub PrintPDF(PDFname)


    Dim PdfJob   As Object

    Dim ws       As Worksheet

    Dim closepdf As Boolean

    Dim i        As Integer

    Dim counter  As Integer

    Dim rc       As Long
    
    Dim DefaultPrinter As String
    
    Dim Source As String

    ' termine toute tache en cours si PDFCreator est encore en exécution
   ' killtask ("PDFCreator.exe")
    KillProcess ("PDFCreator.exe")
  
    'Set PdfJob = New PDFCreator.clsPDFCreator
    Set PdfJob = CreateObject("PDFCreator.clsPDFCreator")

    With PdfJob
  
    'closepdf = PdfJob.cStart("/NoProcessingAtStartup")
        If .cStart("/NoProcessingAtStartup") = False Then
        MsgBox "DESOLE... impossible d'initialiser PDF Creator..." & vbCr & _
        "Veuillez voir le problème et relancer l'opération plus tard S.V.P..."
          Exit Sub
        End If

        .cDefaultPrinter = True

        .cOption("UseAutosave") = 1

        .cOption("UseAutosaveDirectory") = 1

        .cOption("AutosaveDirectory") = ActiveWorkbook.Path & "\"
        '  MsgBox (GetFilePath(PDFname))
        '  MsgBox (GetFileName(PDFname))
        .cOption("AutosaveFilename") = "COMP"

        .cOption("AutosaveFormat") = 0 ' 0 = PDF
    
    

        DefaultPrinter = .cDefaultPrinter
    
        .cDefaultPrinter = "PDFCreator"
        .cCombineAll
        .cClearCache
   End With
   
'   Source = ActiveWorkbook.Path & "\COMP.xls"
'   If FileExists(Source) Then
'
'                    rc = ShellExecute(0, "print", Source, "", "", 0)
'                     Do Until PdfJob.cCountOfPrintjobs = 1 'attend d'avoir un fichier en attente
'                      DoEvents
'                     Loop
'                     If 1 < rc And rc <= 32 Then
'
'                        Debug.Print "Un truc c'est mal passé avec l'impression."
'
'
'                    End If
'    End If
'
    ActiveWorkbook.PrintOut Copies:=1, Collate:=True
   
   
      ' arrete la tache d'impression (pour éviter la création
   'automatique du premier fichier)
  '  PdfJob.cPrinterStop = True

  PdfJob.cPrinterStop = False

    Do Until PdfJob.cCountOfPrintjobs = 0 'attend que l'impression soit dans la file d'attente de pdf créator
        DoEvents
    Loop
   
   ' restauration de l'état
   
   '*************************
   '*************************
   With PdfJob
     .cDefaultPrinter = DefaultPrinter
     Sleep 3000
     .cClearCache
     Sleep 3000
     .cClose
     ' Attendre jusqu'à ce que PDFCreator soit supprimé de la mémoire
     Sleep 3000

   End With
   
  '**********************
  '*********************
   Set PdfJob = Nothing
 
End Sub

Sub New_PrintComp()
Attribute New_PrintComp.VB_ProcData.VB_Invoke_Func = "A\n14"
'Il faut trouver les indices des dernières lignes et colonnes
Dim Destination As String
Dim DEQ As String
Dim fincom As String
Dim i, j, k As Integer
'***********************************************************
'***********************************************************
'Tester s'il existe un doublon dans les champs des caractéristiques

    Dim Plage As Range
    Dim Cel As Range
 
    With ActiveWorkbook.Sheets(1)
 
    'en colonne "B" à partir de B19
       Set Plage = .Range(.Cells(19, 2), .Cells(.Rows.Count, 2).End(xlUp))
 
    End With
    
    'boucle la plage de la feuille "Compte" et cherche chaque valeur
   'en correspondance exacte dans la plage de la feuille "Source"
   For Each Cel In Plage
 
        If Application.CountIf(Plage, Cel.Value) > 1 Then
 
            MsgBox "Attention, la valeur '" & Cel.Value & "' est en doublon," _
                   & " veuillez éliminer manuellement le double situé en '" & Cel.Address(0, 0) _
                   & "' avant de pouvoir exporter les données !"
 
            Cel.Interior.ColorIndex = 13
 
        End If
 
    Next Cel


'***********************************************************
'***********************************************************
    Columns("E:G").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll Down:=-21
    Range("E13").Select
    ActiveWindow.SmallScroll Down:=6
    
DEQ = GetFileName(ActiveWorkbook.Path)
MsgBox (DEQ)


MsgBox (ActiveWorkbook.Sheets(1).Cells(1, 3).Value)
If ActiveWorkbook.Sheets(1).Cells(1, 3).Value = DEQ Or ActiveWorkbook.Sheets(1).Cells(1, 4).Value = DEQ Then
PrintPDF (Destination)
MsgBox ("OK")
Else
MsgBox ("Attention !! il y a une erreur dans le numero du DEQ!")
End If
'***********************************************


'Tester si une case vide dans le comparatif

       
       fincom = ActiveWorkbook.Sheets(1).Columns(2).Cells.Find("*", , , , , xlPrevious).Row

If fincom > 18 Then

     For i = 18 To fincom + 1

                If ActiveWorkbook.Sheets(1).Cells(i, 2).Value = "" Then
                    Cells(i, 2).Select
                    Cells(i, 2).FormatConditions.Delete
                    Cells(i, 2).Interior.Color = RGB(255, 0, 0) ' Rouge
        
                End If

                
                If ActiveWorkbook.Sheets(1).Cells(i, 3).Value = "" Then
                    Cells(i, 3).Select
                    Cells(i, 3).FormatConditions.Delete
                    Cells(i, 3).Interior.Color = RGB(255, 0, 0) ' Rouge
        
                End If
      
        
                If ActiveWorkbook.Sheets(1).Cells(i, 4).Value = "" Then
                   Cells(i, 4).Select
                   Cells(i, 4).FormatConditions.Delete
                   Cells(i, 4).Interior.Color = RGB(255, 0, 0) ' Rouge
        
                End If

    Next i



End If
     
       
'************************************************

 
End Sub

Sub killtask(sappname As String)
' /** j'ai pris le code tel quel... pas de commentaire/
Dim oProclist As Object
Dim oWMI As Object
Dim oproc As Object
Set oWMI = GetObject("winmgmts:")
If IsNull(oWMI) = False Then
    Set oProclist = oWMI.InstancesOf("win32_process")
    For Each oproc In oProclist
        If UCase(oproc.Name) = UCase(sappname) Then
            oproc.Terminate (0)
        End If
    Next oproc
Else
    MsgBox "Killing """ & sappname & _
    """ - Can't create WMI Object.", _
    vbOKOnly + vbCritical, "CloseAPP_B"
End If
Set oProclist = Nothing
Set oWMI = Nothing
End Sub
' /** 2ème alternatives pour fermer le process en cours (suite à une erreur avec killtask)
Public Function KillProcess(ByVal ProcessName As String) As Boolean
On Error GoTo err
    Dim svc As Object
    Dim sQuery As String
    Dim oproc
    Set svc = GetObject("winmgmts:root\cimv2")
    sQuery = "select * from win32_process where name='" & ProcessName & "'"
    For Each oproc In svc.execquery(sQuery)
        oproc.Terminate
    Next
    Set svc = Nothing
err:
Set svc = Nothing
'MsgBox err.Description & " " & err.Number & " " & err.Source
End Function


Function GetFileName(ByVal FileName As String) As String

    Dim i As Integer
    
    

 
    GetFileName = FileName

    For i = Len(FileName) To 1 Step -1

        If Mid(FileName, i, 1) = ":" Or Mid(FileName, i, 1) = "\" Then

            GetFileName = Mid(FileName, i + 1)

            Exit For

        End If

    Next

End Function




