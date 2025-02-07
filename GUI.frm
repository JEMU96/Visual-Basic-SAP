VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GUI 
   Caption         =   "Selector de Pagos"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "GUI.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    MsgBox "Se cierra el Programa sin Aplicar Pago.", , "Cancelado por Usuario"
    Me.Hide
    AutoPayments.Close
End Sub
Private Sub UserForm_Activate()
    ComboBox1.AddItem "Pagares"
    ComboBox1.AddItem "Confirming"
    ComboBox1.AddItem "Otros Pagos"
        
    ListBoxPag.AddItem "Pagaré Con Subsidiarias"
    ListBoxPag.AddItem "Pagarés Único Cliente"
    
    ListBoxConf.AddItem "Confirming"
     
    ListBoxOtros.AddItem "Movimientos Bancarios"
    ListBoxOtros.AddItem "Pagos Diarios"

    ' Hide all list boxes initially
    ListBoxPag.Visible = False
    ListBoxConf.Visible = False
    ListBoxOtros.Visible = False
End Sub

Private Sub ComboBox1_Change()
    ' Hide all list boxes
    ListBoxPag.Visible = False
    ListBoxConf.Visible = False
    ListBoxOtros.Visible = False
    
    ' Show the selected list box
    Select Case ComboBox1.Value
        Case "Pagares"
            ListBoxPag.Visible = True
        Case "Confirming"
            ListBoxConf.Visible = True
        Case "Otros Pagos"
            ListBoxOtros.Visible = True
    End Select
End Sub

Private Sub cmdContinue_Click()
    ' Check which list box is visible and call the appropriate macro based on the selected item
    If ListBoxPag.Visible Then
        Select Case ListBoxPag.Value
            Case "Pagaré Con Subsidiarias"
                ModulePag.PromissoryNote_Multiple_Clients
                Me.Hide
            Case "Pagarés Único Cliente"
                ModulePag.Same_Client_Mutiple_PromissoryNotes
                Me.Hide
        End Select
    ElseIf ListBoxConf.Visible Then
        Select Case ListBoxConf.Value
            Case "Confirming"
                ModuleConf.ConfirminOneClient
        End Select
    ElseIf ListBoxOtros.Visible Then
        Select Case ListBoxOtros.Value
            Case "Movimientos Bancarios"
                Call ModuleDiario.BankFile
                Me.Hide
            Case "Pagos Diarios"
                Call ModuleDiario.DailyPayments
                Me.Hide
        End Select
    End If
End Sub
Sub ShowForm()
    UserForm1.Show
End Sub

Private Sub UserForm_Terminate()
    Program.Close
End Sub
