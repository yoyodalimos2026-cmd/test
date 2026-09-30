Attribute VB_Name = "Module3"
Sub CréationRépertoire()
Dim l, i, comp, col As Integer
Dim Message, Rep As String
'il faut ajouter au projet la référence  "Ms Scripting runtime"
Dim fs As Scripting.FileSystemObject
Set fs = New Scripting.FileSystemObject
Message = " Les dossier :" & vbCr
comp = 0
col = 0
l = Range("A65536").End(xlUp).Row
For i = 2 To l
Rep = "E:\Validation des équivalents\En_cours\" & Cells(i, 1).Value
col = col + 1
  ' test si le repertoire existe
  If fs.FolderExists(Rep) Then
 Message = Message & Cells(i, 1).Value & vbCr
Else
MkDir (Rep)
comp = comp + 1
  End If
'End If
Next i
MsgBox (Message & "Existe déja !" & vbCr & comp & "/" & col & " Dossiers ont été créés !")
'MsgBox (Message & "Existe déja !" & comp & "/" & col & " Dossiers ont été créés !")


  Set fs = Nothing



End Sub
