Attribute VB_Name = "ModuleSAP"
'Public VARs
Public Ses As Object
Public Continue As Boolean
Public MainAccount
Public VTO




Public Declare PtrSafe Function MessageBox _
    Lib "User32" Alias "MessageBoxA" _
       (ByVal hWnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal wType As Long) _
    As Long
'-Begin-----------------------------------------------------------------
Dim gColl() As String
Dim nColl() As String
Dim tColl() As String
Dim typeColl() As String
Dim j As Integer
Sub GetAll(Obj) '---------------------------------------------
'-
'- Recursively called sub routine to get the IDs of all UI elements
'-
'-----------------------------------------------------------------------

  Dim cntObj As Integer
  Dim i As Integer
  Dim Child As Object
  Dim a As String

  On Error Resume Next
  cntObj = Obj.Children.Count()
  If cntObj > 0 Then
    For i = 0 To cntObj - 1
      Set Child = Obj.Children.Item(CLng(i))
      GetAll Child
      ReDim Preserve gColl(j)
      ReDim Preserve nColl(j)
      ReDim Preserve tColl(j)
      ReDim Preserve typeColl(j)
      gColl(j) = CStr(Child.ID)
      nColl(j) = CStr(Child.Name)
      typeColl(j) = CStr(Child.Type)
'      If typeColl(j) = "GuiTitlebar" Then
'        pause
'      End If
      If typeColl(j) = "GuiButton" Then
        tColl(j) = CStr(Child.ToolTip)
      Else
        tColl(j) = CStr(Child.Text)
      End If
      j = j + 1
    Next
  End If
  On Error GoTo 0
End Sub
Function ItemsFoundSAP(Ses)
'Return a Dictionary of all the items found in SAP
    titl = ChkWindow()
    Do While InStr(1, titl, "Procesar partidas abiertas") = 0
        Ses.findById("wnd[0]/tbar[1]/btn[16]").press
        titl = ChkWindow()
    Loop
    nPA = CInt(Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-ANZPO").Text) ' Number of Items found on SAP
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    n = dict.Count
    ii = 1
    Do While nPA <> n
        GetAll Ses
        For i = LBound(gColl) To UBound(gColl)
        On Error Resume Next
            If nColl(i) = "RFOPS_DK-XBLNR" Then
                dict.Add tColl(i), tColl(i)
            End If
        Next
        n = dict.Count
        Ses.findById("wnd[0]").sendVKey 82
    Loop
    ' Return the dictionary
    Set ItemsFoundSAP = dict
    Ses.findById("wnd[0]/tbar[1]/btn[14]").press 'Navigate to Accounting entry sumary

End Function
Function ConnectToSAP():
'Connect to the SAP GUI and return the first session (child) found
    Dim WScript As Object
    On Error Resume Next
    If Not IsObject(SAPapp) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAPapp = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(SAPConnection) Then
        Set SAPConnection = SAPapp.Children(0)
    End If
    If Not IsObject(SAPsession) Then
        Set SAPsession = SAPConnection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject SAPsession, "on"
        WScript.ConnectObject SAPapp, "on"
    End If
    Set ConnectToSAP = SAPsession
    On Error GoTo 0
End Function
Function ChkWindow()
'Returns the title of the window user is currently in
    titl = Ses.findById("wnd[0]").Text
    ChkWindow = titl
End Function
Function ChkSBar()
'Returns the title of the current status bar, which can be EMPTY ""
    titl = Ses.findById("wnd[0]/sbar").Text
    Etype = Ses.findById("wnd[0]/sbar").Messagetype
    If Etype = "E" Or Etype = "W" Then
        ContConfirmation = MsgBox(titl, vbOKCancel, "Ojo!!")
        If ContConfirmation = vbOK Then
            ChkSBar = titl
            Exit Function
        ElseIf ContConfirmation = vbCancel Then
            MsgBox "Cancelado por el usuario", , "Cancelaci n"
            Continue = False
            Exit Function
        End If
    End If
    ChkSBar = titl
End Function
Function CallTransaction(Transaccion As String)
'Run a specified transaction in SAP
    titl = ChkWindow()
    If InStr(1, titl, "900 SAP Easy Access") = 0 Then
        BackToMain
    End If
    Ses.findById("wnd[0]/tbar[0]/okcd").Text = Transaccion
    Ses.findById("wnd[0]").sendVKey 0
End Function
Function BackToMain()
'Return to the main menu without applying any changes
    Ses.findById("wnd[0]/tbar[0]/okcd").Text = "/n000"
    Ses.findById("wnd[0]").sendVKey 0
End Function
Function BatchInput(Template As String)
'Call the Transaction Batch Input to load a Template.
    titl = ChkWindow()
    If InStr(1, titl, "900 SAP Easy Access") = 0 Then
        BackToMain
    End If
    CallTransaction ("Z2S_K0021")
    Ses.findById("wnd[0]/usr/radP_CALLT").Select ' Select to Call a Transaction from the Template
    Ses.findById("wnd[0]/usr/ctxtP_FILE").Text = Template
    Ses.findById("wnd[0]/tbar[1]/btn[8]").press ' Run
    Ses.findById("wnd[1]/usr/btnSPOP-OPTION1").press ' Press Continue on Pop-up Window
    titl = ChkSBar()
    If titl = "Por favor, seleccione primero las partidas." Then
        ask = MessageBox(&H0, "No se han seleccionado PAs." & vbCrLf & _
        " Es un pago sin PA?", "Confirmaci n", vbYesNo)
        If ask = vbNo Then
            MsgBox "Se cancela el proceso.", vbOKOnly, "Cancelaci n)"
            Continue = False
            BackToMain
        ElseIf ask = vbYes Then
            Ses.findById("wnd[0]").sendVKey 12 ' Cancel Select PA's and go to enter payments in the account
        End If
    Else
        Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
        Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
        Ses.findById("wnd[0]/tbar[1]/btn[14]").press 'Monta ita
    End If
End Function
Function SAPData()
    'Return a Dict with some important Amounts from SAP
    titl = ChkWindow()
    Do While InStr(1, titl, "Procesar partidas abiertas") = 0
        Ses.findById("wnd[0]/tbar[1]/btn[16]").press
        titl = ChkWindow()
    Loop
    Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
    Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
    NumPAs = Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-ANZPO").Text ' Number of Items found on SAP
    ImpPAsSAP = Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-NETTO").Text ' Sum of the values of  the items found
    ImpDifSAP = Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB").Text ' Total amount that must be enter to apply the accounting entry
    ImpUSER = Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-BETRG").Text ' Total amount entered by the user
    'Format String to Number
    NumPAs = FormatNumber(NumPAs, 2)
    ImpPAsSAP = FormatNumber(ImpPAsSAP, 2)
    ImpDifSAP = FormatNumber(ImpDifSAP, 2)
    ImpUSER = FormatNumber(ImpUSER, 2)
    'Prepare the result list
    Dim ResultData As Object
    Set ResultData = CreateObject("Scripting.Dictionary")
    ResultData.Add "NumPAs", NumPAs
    ResultData.Add "ImpPAsSAP", ImpPAsSAP
    ResultData.Add "ImpDifSAP", ImpDifSAP
    ResultData.Add "ImpUSER", ImpUSER
    Set SAPData = ResultData
    Ses.findById("wnd[0]/tbar[1]/btn[14]").press 'Navigate to Accounting entry sumary
End Function
Function NewEntry(Code As String, Account, Optional CME = "", Optional DocDate = "") 'Finish
    'SAP ID Fields
    CodeField = "wnd[0]/usr/ctxtRF05A-NEWBS"
    AccountNumbField = "wnd[0]/usr/ctxtRF05A-NEWKO"
    CMEField = "wnd[0]/usr/ctxtRF05A-NEWUM"
    DocDateField = "wnd[0]/usr/ctxtBKPF-BLDAT"
    titl = ChkWindow()
    If InStr(1, titl, "Liquidar compensaci n: Datos cabecera") > 0 Then
        'Select "Traslados con compensaci n"
        Ses.findById("wnd[0]/usr/sub:SAPMF05A:0122/radRF05A-XPOS1[3,0]").Select
        If DocDate <> "" Then
            Ses.findById(DocDateField).Text = DocDate
        Else
            DocDate = ModuleAux.AskUserDate()
            Ses.findById(DocDateField).Text = DocDate
        End If
    End If
    If (Code = "09" Or Code = "19") And CME = "" Then
        MsgBox "Clave CME no introducida"
        Continue = False
        BackToMain
        Exit Function
    ElseIf (Code = "90" Or Code = "91") And (UCase(CME) = "Z" Or UCase(CME) = "K") Then
        Continue = True
    ElseIf (Code = "60" Or Code = "61") And CME = "" Then
        Continue = True
    ElseIf (Code = "40" Or Code = "50") And CME = "" Then
        Continue = True
    ElseIf (Code = "26" Or Code = "36") And CME = "" Then
        Continue = True
    Else
        MsgBox "Esas claves no estan incluidas en el programa"
        Continue = False
        BackToMain
        Exit Function
    End If
    'Set text
    Ses.findById(CodeField).Text = Code
    Ses.findById(AccountNumbField).Text = Account
    Ses.findById(CMEField).Text = CME

    Ses.findById("wnd[0]").sendVKey 0 ' Returnkey
    sbar = ChkSBar()
    If sbar <> "" Then
        ask = MsgBox(sbar, vbOKCancel, "Ojo!")
    End If
End Function
Function NewEntryAddData(Amount, DueTo, Optional Comentary = "", Optional PaymentMethod = "", Optional Asignation = "", Optional CECO = "")
    On Error Resume Next
    Division = "AAAA" ' Buisnes Division "Sensitive Data"
    If Asignation = "" Then
        Asignation = Right(DueTo, 4) & Mid(DueTo, 4, 2) & Left(DueTo, 2)
    End If
    'SAP Fields
    AmountField = "wnd[0]/usr/txtBSEG-WRBTR"
    DueToField = "wnd[0]/usr/ctxtBSEG-ZFBDT"
    PaymentMethodFiled = "wnd[0]/usr/ctxtBSEG-ZLSCH"
    AsignationField = "wnd[0]/usr/txtBSEG-ZUONR"
    ComentaryFiled = "wnd[0]/usr/ctxtBSEG-SGTXT"
    
    'Posible Division Fields
    DivisionField = "wnd[0]/usr/ctxtBSEG-GSBER"
    DivisionField2 = "wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-GSBER"
    DivisionField3 = "wnd[0]/usr/subBLOCK:SAPLKACB:1010/ctxtCOBL-GSBER"
    DivisionField4 = "wnd[1]/usr/ctxtCOBL-GSBER"
    
    'Posible CECO Fields
    CECOField = "wnd[0]/usr/subBLOCK:SAPLKACB:1010/ctxtCOBL-KOSTL"
    CECOField2 = "wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-KOSTL"
    CECOField3 = "wnd[1]/usr/ctxtCOBL-KOSTL"
    
    'Enter Data fix Fields
    Ses.findById(AmountField).Text = Amount
    Ses.findById(DueToField).Text = DueTo
    Ses.findById(AsignationField).Text = Asignation
    Do While Comentary = ""
        Comentary = InputBox("Introduce un comentario para el apunte", "Comentario")
    Loop
    Ses.findById(ComentaryFiled).Text = Comentary
    If PaymentMethod = "" Then
        ask = MessageBox(&H0, " El apunte tiene v a de pago?", "Confirmaci n de V a de Pago", vbYesNo)
        If ask = vbYes Then
            Do While PaymentMethod = ""
                PaymentMethod = InputBox("Introduce la v a de pago", "Confirmaci n de V a de Pago")
                If PaymentMethod <> "2" And PaymentMethod <> "T" And PaymentMethod <> "R" And PaymentMethod <> "3" Then
                    MsgBox "V a de pago no valida." & vbCrLf & "Por favor introduce una v a de pago valida"
                    PaymentMethod = ""
                End If
            Loop
        End If
    ElseIf PaymentMethod = -1 Then
        PaymentMethod = ""
    End If
    Ses.findById(PaymentMethodFiled).Text = PaymentMethod
    Ses.findById("wnd[0]").sendVKey 0 ' Returnkey
    sbar = ChkSBar()
    If sbar <> "" Then
        ask = MsgBox(sbar, vbOKCancel, "Ojo!")
        Ses.findById("wnd[0]").sendVKey 0 ' Returnkey
    End If
    'Enter Multiple posibilities fields data
    Ses.findById(DivisionField).Text = Division
    Ses.findById(DivisionField2).Text = Division
    Ses.findById(DivisionField3).Text = Division
    Ses.findById(DivisionField4).Text = Division
    If CECO <> "" Then
        Ses.findById(CECOField).Text = CECO
        Ses.findById(CECOField2).Text = CECO
        Ses.findById(CECOField3).Text = CECO

    End If
    Ses.findById("wnd[0]").sendVKey 0 ' Returnkey

    sbar = ChkSBar()
    If sbar <> "" Then
        ask = MsgBox(sbar, vbOKCancel, "Ojo!")
    End If
    Ses.findById("wnd[0]/tbar[1]/btn[14]").press 'Navigate to Accounting entry sumary
    Ses.findById("wnd[0]").sendVKey 0 ' Returnkey
End Function
Function Simulate()
'Return List of positions to fill in data in each one
    titl = ChkWindow()
    Do While InStr(1, titl, "Visualizar Resumen") = 0
        Ses.findById("wnd[0]/tbar[1]/btn[14]").press
        titl = ChkWindow()
    Loop
    PosIni = Ses.findById("wnd[0]/usr/txtRF05A-ANZAZ").Text
    Ses.findById("wnd[0]/mbar/menu[0]/menu[3]").Select 'simular
    titl = ChkSBar()
    If titl = "La diferencia es demasiado grande para una compensaci n" Then
        DifSAP = Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB").Text     ' Diference after enter the data
        DifSAP = FormatNumber(DifSAP, 2)
        response = MsgBox("La diferencia es: " & DifSAP & vbCrLf & _
                " Quiere redondearla o dejarla en la cuenta?", vbYesNo, "Diferencia encontrada")
        If resposne = vbNo Then
            MsgBox "Se cancela el proceso.", vbCancel
            Continue = False
            BackToMain
            Exit Function
        ElseIf response = vbYes Then
            Ses.findById("wnd[0]/tbar[1]/btn[14]").press 'Navigate to Accounting entry sumary
            If DifSAP < -1 Or DifSAP > 1 Then
                MsgBox " a cuenta"
                'llamar ToAccountDif
                ToAccountDif DifSAP
            Else
                DifConfirmation.DifMSG DifSAP
            End If
            
            Ses.findById("wnd[0]/mbar/menu[0]/menu[3]").Select 'simular
            Continue = True
        End If
        
    End If
    PosEnd = Ses.findById("wnd[0]/usr/txtRF05A-ANZAZ").Text
    Dim Positions(1)
    Positions(0) = PosIni
    Positions(1) = PosEnd
    Simulate = Positions
End Function
Function EnterPosition(Position)
    Ses.findById("wnd[0]/usr/txtRF05A-ANZAZ").SetFocus
    Ses.findById("wnd[0]").sendVKey 2
    Ses.findById("wnd[1]/usr/txt*BSEG-BUZEI").Text = Position
    Ses.findById("wnd[1]/tbar[0]/btn[13]").press
End Function
Function ToAccountDif(DifSAP)
    If DifSAP < 0 Then
        DifSAP = DifSAP * -1
        NewEntry "60", MainAccount
        NewEntryAddData DifSAP, VTO
    ElseIf DifSAP > 0 Then
        NewEntry "61", MainAccount
        NewEntryAddData DifSAP, VTO
    Else
        MsgBox "Error"
    End If
    'MsgBox "La diferencia se deja en la cuenta del cliente " & DifSAP
End Function
Function RoundDif(DifSAP)
CECO = "AAA12158541" ' Buisnes Round Cost Center "Sensitive Data"
    If DifSAP < 0 Then
        DifSAP = DifSAP * -1
        NewEntry "40", "123456789"
        NewEntryAddData DifSAP, VTO, , -1, , CECO
    ElseIf DifSAP > 0 Then
        NewEntry "50", "123456789"
        NewEntryAddData DifSAP, VTO, , -1
    Else
        MsgBox "Error"
    End If
       
    'MsgBox "La diferencia se redondea " & DifSAP
End Function
Function EnterAJD(Amount, Asignation, Comentary)
    If Amount < 0 Then
        Amount = Amount * -1
    End If
    CECO = "ZZZZ12158541" ' Buisnes Bank Expenses Cost Center "Sensitive Data"
    NewEntry "40", "55555555"
    NewEntryAddData Amount, VTO, Comentary, -1, Asignation, CECO
End Function
Function SearchItems(Category, Position, Optional SearchData = "", Optional SocietyCode = "", Optional Account = "", Optional AditionalSearchData = "") 'Revisar
    'Ning. Pos= 0
    'Importe Pos= 1
    'N  documento Pos= 2
    'Fe.contabilizaci n Pos= 3
    'Referencia a factura Pos= 4
    'Referencia Pos= 5
    'Clase de documento Pos= 6
    'Indicador impuestos Pos= 7
    'Solicitud acept.L/C Pos= 8
    'Cta.subsidiaria Pos= 9
    'Moneda Pos= 10
    'Clave contabiliz. Pos= 11
    'Fecha de documento Pos= 12
    'Asignaci n Pos= 13
    'Factura Pos= 14
    'Posici n Pos= 15
    'Vencimiento neto Pos= 16
    titl = ChkWindow()
    Do While InStr(1, titl, "Visualizar Resumen") = 0
        ask = Message + Box(&H0, "No esta en la venta apropiada." & vbCrLf & _
        "Vaya a Visualizar Resumen y presione Ok.", "Confirmaci n", vbOKCancel)
        If ask = vbCancel Then
            Continue = False
            BackToMain
            Exit Function
        Else
            titl = ChkWindow()
        End If
    Loop
    SocField = "wnd[0]/usr/ctxtRF05A-AGBUK"
    AccField = "wnd[0]/usr/ctxtRF05A-AGKON"
    CatField = "wnd[0]/usr/ctxtRF05A-AGKOA"
    Ses.findById(CatField).Text = Category
    If SocietyCode <> "" Then Ses.findById(SocField).Text = SocietyCode
    If Account <> "" Then Ses.findById(AccField).Text = Account
    If Postion = 1 Then
        SearchFiled1 = "wnd[0]/usr/sub:SAPMF05A:0730/txtRF05A-VONWT[0,0]"
        SearchField2 = "wnd[0]/usr/sub:SAPMF05A:0730/txtRF05A-BISWT[0,21]"
    ElseIf Position = 16 Then
        SearchFiled1 = "wnd[0]/usr/sub:SAPMF05A:0732/ctxtRF05A-VONDT[0,0]"
        SearchField2 = "wnd[0]/usr/sub:SAPMF05A:0732/ctxtRF05A-BISDT[0,20]"
    ElseIf Position = 5 Then
        SearchFiled1 = "wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]"
        SearchField2 = "wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL02[0,31]"
    Else
        MsgBox "No se ha contemplado esa seleccion", , "Cancelar"
        Continue = False
        Exit Function
    End If
    Ses.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[" & Position & ",0]").Select
    Ses.findById("wnd[0]").sendVKey 0 ' Return
    If SearchData <> "" Then
        Ses.findById("wnd[0]/usr/sub:SAPMF05A:0730/txtRF05A-VONWT[0,0]").Text = SearchData
        Ses.findById("wnd[0]").sendVKey 0 ' Return
        Ses.findById("wnd[0]/tbar[1]/btn[16]").press ' Navigate to Select Items
    End If
    If AditionalSearchData <> "" Then
        Ses.findById("wnd[0]/usr/sub:SAPMF05A:0730/txtRF05A-VONWT[0,0]").Text = AditionalSearchData
        Ses.findById("wnd[0]").sendVKey 0 ' Return
        Ses.findById("wnd[0]/tbar[1]/btn[16]").press ' Navigate to Select Items
    End If
    

    titl = ChkSBar()
    If titl <> "" Then
        MsgBox titl, , "Ojo!"
    End If
    titl = ChkWindow()
    If InStr(1, titl, "Procesar partidas abiertas") = 0 Then
        BackToMain
        Continue = False
    End If
    Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
    Ses.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
End Function
Function GetEntryNumber()
    titl = ChkWindow()
    Do While InStr(1, titl, "Visualizar documento: Acceso") = 0
        ask = MessageBox(&H0, "No esta en la venta apropiada." & vbCrLf & _
        "Vaya a Visualizar Documento y presione Ok.", "Confirmaci n", vbOKCancel)
        If ask = vbCancel Then
            Continue = False
            BackToMain
            Exit Function
        Else
            titl = ChkWindow()
        End If
    Loop
    EntryNumber = Ses.findById("wnd[0]/usr/txtRF05L-BELNR").Text
    GetEntryNumber = EntryNumber
End Function
Function SaveEntry(Path As String)
    titl = ChkWindow()
    Do While InStr(1, titl, "Visualizar documento:Acceso") = 0
        ask = MessageBox(&H0, "No esta en la venta apropiada." & vbCrLf & _
        "Vaya a Visualizar Documento y presione Ok.", "Confirmaci n", vbOKCancel)
        If ask = vbCancel Then
            Continue = False
            BackToMain
            Exit Function
        Else
            titl = ChkWindow()
        End If
    Loop
    EntryNumber = Ses.findById("wnd[0]/usr/txtRF05L-BELNR").Text
    Ses.findById("wnd[0]").sendVKey 0
    titl = ChkSBar()
    Do While titl <> ""
        application.Wait (Now() + TimeValue("00:00:10"))
        Ses.findById("wnd[0]").sendVKey 0
        titl = ChkSBar()
    End If
    Ses.findById("wnd[0]/tbar[0]/btn[86]").press
    Ses.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").SetFocus
    Ses.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").Key = ""
    Ses.findById("wnd[1]/tbar[0]/btn[13]").press
    Ses.findById("wnd[0]").sendVKey 0
    
    ModuleSAP.BackToMain
    ModuleSAP.CallTransaction "SP01"

    Ses.findById("wnd[0]").sendVKey 8
    Ses.findById("wnd[0]/usr/lbl[3,3]").SetFocus
    Ses.findById("wnd[0]").sendVKey 2
    Ses.findById("wnd[0]/usr/txtTSP01_SP0R-RQTITLE").Text = EntryNumber
    Ses.findById("wnd[0]/tbar[1]/btn[8]").press
    Ses.findById("wnd[1]/usr/btnBUTTON_1").press
    Ses.findById("wnd[0]/usr/chk[1,3]").Selected = True
    Ses.findById("wnd[0]/usr/chk[1,3]").SetFocus
    Ses.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[2]").Select
    Ses.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
    Ses.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = EntryNumber & ".pdf"
    Ses.findById("wnd[1]/tbar[0]/btn[0]").press
    
    ModuleSAP.BackToMain
    
End Function
