VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4740
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
Choix_CBO = cboComboBox.Value
UserForm1.Hide
End Sub



Private Sub CommandButton2_Click()
Choix_CBO = "SILICON"
UserForm1.Hide
End Sub



Private Sub UserForm_Initialize()
Dim Onglet_Début, nbre_Onglet, i As Integer
Dim Chemin As String

'ouvrir Fichier Trame

'parcourir les onglets
Onglet_Début = 3
nbre_Onglet = Workbooks(Fichier_trame).Sheets.Count
'Renseigner les donnée du ComboBox
For i = Onglet_Début To nbre_Onglet
    cboComboBox.AddItem Workbooks(Fichier_trame).Sheets(i).Name
Next i

'Workbooks(fichier).Close
End Sub

