VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DifConfirmation 
   Caption         =   "¿Como aplicar la diferencia?"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "DifConfirmation.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "DifConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConfirmationDif
Sub DifMSG(Dif)
'Dif Confirmation Gui
    ConfirmationDif = Dif
    labeltext = "¿Qué hacemos con la diferencia? " & ConfirmationDif ' Spanish
    DifConfirmation.Label1.Caption = labeltext
    DifConfirmation.Show
End Sub

Private Sub CmdRound_Click()
'Apply Dif to Round Dif Account
    DifConfirmation.Hide
    ModuleSAP.RoundDif ConfirmationDif
End Sub

Private Sub CmdToAccount_Click()
'Apply Dif to Client Account
    DifConfirmation.Hide
    ModuleSAP.ToAccountDif ConfirmationDif
End Sub
