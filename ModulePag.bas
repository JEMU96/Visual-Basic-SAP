Attribute VB_Name = "ModulePag"
Sub Same_Client_Mutiple_PromissoryNotes()
    On Error GoTo ErrorHandler
    Continue = True
    'Client Variables
    Client = "12345"
    ClientName = "ClientName"
    'Public Variable
    MainAccount = Client
       
    PaymentDetailPath = ModuleAux.OpenFile("Abre el fichero de Pagares de Alcampo")
    If PaymentDetailPath = False Then Exit Sub
    
    Set PDwb = Workbooks.Open(PaymentDetailPath)
    Set PDws = PDwb.Sheets(1)
    PDws.Unprotect Password:=ClientName
    EndRow = PDws.Cells(Rows.Count, 1).End(xlUp).Row
    If EndRow <= 1 Then
        MsgBox ("Prepare el fichero correctamente." & vbCrLf & "Proceso cancelado.")
        Exit Sub
    End If
    Set Ses = ModuleSAP.ConnectToSAP ' Connet to SAP GUI
    For Each cell In PDws.Range("A2:A" & EndRow)
        looprow = cell.Row
    
        ' Public Variable
        VTO = PDws.Cells(looprow, 7).Value
        AJD = PDws.Cells(looprow, 9).Value
    
        'Get Data from Worksheet
        DocDate = PDws.Cells(looprow, 1).Value
        Amount = PDws.Cells(looprow, 3).Value
        NPayment = PDws.Cells(looprow, 5).Value
        Comentary = "PAG. " & ClientName & " " & NPayment & " VTO. " & VTO

        
        'Set Data to Worksheet
        PDws.Cells(looprow, 2).Value = DocDate
        PDws.Cells(looprow, 4).Value = Format(DocDate, "yyyymmdd")
        PDws.Cells(looprow, 6).Value = Comentary
        PDws.Cells(looprow, 8).Value = Format(VTO, "yyyymmdd")
        PDws.Cells(looprow, 10).Value = Amount - AJD
        'Format Varariables
        AsignationVTO = Format(VTO, "yyyymmdd")
        AsignationDoc = Format(DocDate, "yyymmdd")
        DocDate = Format(DocDate, "dd.mm.yyyy")
        VTO = Format(VTO, "dd.mm.yyyy")
        Amount = FormatNumber(Amount, 2)
        AJD = FormatNumber(AJD, 2)
        AsignationAJD = "G44_04_" & ClientName 'Business AJD Codification "Sensitive Data"
        SearchAmount = FormatNumber(Amount - AJD, 2)

        'Connect to SAP and enter data
        ModuleSAP.CallTransaction "F-04"
        ModuleSAP.NewEntry "90", Client, "Z", DocDate '90 Z Business Promissory Note Entry Codification "Sensitive Data"
        ModuleSAP.NewEntryAddData Amount, VTO, Comentary, -1, AsignationDoc
        ModuleSAP.EnterAJD AJD, AsignationAJD
        ModuleSAP.SearchItems "D", 1, SearchAmount, , Client 'Position 1 = Amount
        If Continue = False Then
            MsgBox "Se cancela el proceso", vbOKOnly, "Cancelacion"
            ModuleSAP.BackToMain
            Exit Sub
        End If
        Set List = ModuleSAP.SAPData
        If List("ImpDifSAP") <> 0 Then DifConfirmation.DifMSG List("ImpDifSAP")
        Positions = ModuleSAP.Simulate()
        For i = Positions(0) + 1 To Positions(1)
            ModuleSAP.EnterPosition i
            ModuleSAP.NewEntryAddData 0, 0, Comentary, -1, AsignationVTO
        Next
        ModuleAux.SaveConfirmation
        If Continue = False Then GoTo ContinueLoop
        ModuleSAP.CallTransaction "FB03"
        EntryNumber = ModuleSAP.GetEntryNumber
        PDws.Cells(looprow, 11).Value = EntryNumber
ContinueLoop:
    Next
    PDws.Protect Password:=ClientName
    PDwb.Save
    PDwb.Close
    Exit Sub
ErrorHandler:
    If Err.Number = 6 Then
        Resume Next
        Err.Clear
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Sub
Sub PromissoryNote_Multiple_Clients()

    On Error GoTo ErrorHandler
    Continue = True
    'Client Variables
    Client = "1234" 'Sensitive Data
    ClientName = "ClientName" 'Sensitive Data
    'Public Variable
    MainAccount = Client
    
    'Clients Dictoniary
    Dim Clients As Object
    Set Clients = CreateObject("Scripting.Dictionary")
    Clients.RemoveAll
    'Sensitive Data
    Clients.Add "ClientName", "12345"
    Clients.Add "ClientName_Subsidiary1", "11111"
    Clients.Add "ClientName_Subsidiary2", "22222"
    Clients.Add "ClientName_Subsidiary3", "33333"
    Clients.Add "ClientName_Subsidiary4", "44444"
    '.....
    
    Dim dicInvoices As Object
    Set dicInvoices = CreateObject("Scripting.Dictionary")
        
    Dim dicFusis As Object
    Set dicFusis = CreateObject("Scripting.Dictionary")
    
    
    Dim PDwb As Workbook
    Dim PDws As Worksheet
    DetailPath = ModuleAux.OpenFile("Abre la realcion de pago")
    If DetailPath = False Then Exit Sub
    Set PDwb = Workbooks.Open(Path)
    Set PDws = PDwb.Sheets(1)
    EndRow = PDws.Cells(Rows.Count, 1).End(xlUp).Row
    IniRow = 5
    
    Dim Template As String
    Template = "C:\servername\Template.xlsx" 'Auto select the template from the Business Server
    'Tmeplate = ModuleAux.OpenFile("Abre la Plantilla Call Transaction") ' Ask User to open Template
    Dim Tempwb As Workbook
    Dim Tempws As Worksheet
    Set Tempwb = Workbooks.Open(Template)
    Set Tempws = Tempwb.Sheets(1)
    
    PDws.Cells(2, 4).Clear
    UserAmount = ModuleAux.AskUserNumber("Introduce el total del Pagaré")
    PDAmount = FormatNumber(application.Sum(PDws.Columns("D:D")), 2)
    
    If UserAmount <> PDAmount Then
        MsgBox "El importe no cuadra. Se cancela el proceso", , "Cancelación"
        Exit Sub
    End If
    

    TempIniRow = 10
    TempEndRow = Tempws.Cells(Rows.Count, 4).End(xlUp).Row
    If TempIniRow < TempEndRow Then
        Tempws.Range("D10:D" & TempEndRow & "").Clear
    End If
    PDws.Range("B:B").Replace what:="-", Replacement:="", LookAt:=xlPart
    
    VTO = PDws.Cells(2, 7).Value
    AsignationVTO = Format(VTO, "yyyymmdd")
    VTO = Format(VTO, "dd.mm.yyyy")
    
    DocDate = Date
    AsignationDoc = Format(DocDate, "yyyymmdd")
    DocDate = Format(DocDate, "dd.mm.yyyy")
    NPayment = PDws.Cells(2, 2).Value
    Invoinces = FormatNumber(0, 2)
    
    ChequeComentary = "PAG. " & ClientName & " " & NPayment & " VTO. " & VTO
    CargosComentary = "TOTAL CARGOS " & ClientName & " " & NPayment & " VTO. " & VTO
    AbonosComentary = "TOTAL ABONOS " & ClientName & " " & NPayment & " VTO. " & VTO
    AJDComentary = "GASTOS AJD " & ClientName & " " & NPayment & " VTO. " & VTO

    
    CargoMain = FormatNumber(0, 2)
    AbonoMain = FormatNumber(0, 2)
    
    CargoSub1 = FormatNumber(0, 2)
    AbonoSub1 = FormatNumber(0, 2)
    
    CargoSub2 = FormatNumber(0, 2)
    AbonoSub2 = FormatNumber(0, 2)
    
    CargoSub3 = FormatNumber(0, 2)
    AbonoSub3 = FormatNumber(0, 2)
    
    AbonoSub4 = FormatNumber(0, 2)
    CargoSub4 = FormatNumber(0, 2)
    
    For i = IniRow To EndRow
        DocType = PDws.Cells(i, 1).Value
        Ref = PDws.Cells(i, 2).Value
        Fus = Left(PDws.Cells(i, 2).Value, 1)
        Sdd = UCase(PDws.Cells(i, 9).Value)
        Value = PDws.Cells(i, 4).Value
        
        If DocType = "Factura" Then
            Tempws.Cells(TempIniRow, 4).Value = PDws.Cells(i, 2).Value
            TempIniRow = TempIniRow + 1
            Invoinces = Invoinces + PDws.Cells(i, 4).Value
            dicInvoices.Add Ref, Value
        ElseIf DocType = "Abono" Then
            If Fus = "F" Then
                dicFusis.Add i, i
            ElseIf InStr(1, Sdd, "ClientName") > 0 Then
                AbonoMain = AbonoMain + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary1") > 0 Then
                AbonoSub1 = AbonoSub1 + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary2") > 0 Then
                AbonoSub2 = AbonoSub2 + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary4") > 0 Then
                AbonoSub4 = AbonoSub4 + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary3") > 0 Then
                AbonoSub3 = AbonoSub3 + PDws.Cells(i, 4).Value
            End If
        ElseIf DocType = "Cargo" Then
            If Fus = "F" Then
                dicFusis.Add i, i
            ElseIf InStr(1, Sdd, "ClientName") > 0 Then
                CargoMain = CargoMain + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary1") > 0 Then
                CargoSub1 = CargoSub1 + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary2") > 0 Then
                CargoSub2 = CargoSub2 + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary3") > 0 Then
                CargoSub3 = CargoSub3 + PDws.Cells(i, 4).Value
            ElseIf InStr(1, Sdd, "ClientName_Subsidiary4") > 0 Then
                CargoSub4 = CargoSub4 + PDws.Cells(i, 4).Value
            End If
        End If
    Next i
    Tempws.Cells(2, 5).Value = DocDate
    Tempws.Cells(2, 7).Value = DocDate
    Tempwb.Save
    Tempwb.Close
    
    Dim DicCargos As Object
    Set DicCargos = CreateObject("Scripting.Dictionary")
    Dim DicAbonos As Object
    Set DicAbonos = CreateObject("Scripting.Dictionary")
    
    DicCargos.Add "ClientName", CargoMain
    DicCargos.Add "ClientName_Subsidiary1", CargoSub1
    DicCargos.Add "ClientName_Subsidiary3", CargoSub3
    DicCargos.Add "ClientName_Subsidiary2", CargoSub2
    DicCargos.Add "ClientName_Subsidiary4", CargoSub4
    
    DicAbonos.Add "ClientName", AbonoMain
    DicAbonos.Add "ClientName_Subsidiary1", AbonoSub1
    DicAbonos.Add "ClientName_Subsidiary3", AbonoSub3
    DicAbonos.Add "ClientName_Subsidiary2", AbonoSub2
    DicAbonos.Add "ClientName_Subsidiary4", AbonoSub4
    
    Set Ses = ModuleSAP.ConnectToSAP ' Connet to SAP GUI
    ModuleSAP.BatchInput Template
    ModuleSAP.NewEntry "90", Client, "Z", DocDate '90 Z Business Promissory Note Entry Codification "Sensitive Data"
    ModuleSAP.NewEntryAddData UserAmount, VTO, ChequeComentary, -1, AsignationDoc
    
    For Each ClientName In Clients.keys
        Cargo = DicCargos(ClientName)
        Abono = DicAbonos(ClientName)
        ClientCode = Clients(ClientName)
        If Cargo <> 0 Then
            Cargo = Cargo * -1
            ModuleSAP.NewEntry "06", ClientCode
            ModuleSAP.NewEntryAddData Cargo, VTO, CargosComentary, -1, AsignationVTO
        End If
        If Abono <> 0 Then
            ModuleSAP.NewEntry "16", ClientCode
            ModuleSAP.NewEntryAddData Abono, VTO, AbonosComentary, -1, AsignationVTO
        End If
    Next ClientName
    For Each FusiRow In dicFusis.keys
        FusiType = PDws.Cells(FusiRow, 1).Value
        FusiValue = FormatNumber(PDws.Cells(FusiRow, 4).Value, 2)
        FusiName = PDws.Cells(FusiRow, 2).Value
        FusiSodd = UCase(PDws.Cells(FusiRow, 9).Value)
        FusiComentary = "CARGO " & FusiName & " REPERCUTIR EDITORES"
        For Each ClientName In Clients.keys
            If InStr(1, FusiSodd, ClientName) > 0 Then
                ClientCode = Clients(ClientName)
                If FusiType = "Abono" Then
                    ModuleSAP.NewEntry "61", Client '61 Business Debit Entry Codification "Sensitive Data"
                    ModuleSAP.NewEntryAddData FusiValue, VTO, FusiComentary, -1, AsignationVTO
                ElseIf FusiType = "Cargo" Then
                    FusiValue = FusiValue * -1
                    ModuleSAP.NewEntry "60", Client '60 Business Credit Entry Codification "Sensitive Data"
                    ModuleSAP.NewEntryAddData FusiValue, VTO, FusiComentary, -1, AsignationVTO
                End If
            End If
        Next ClientName
    Next FusiRow
    Set ResultData = SAPData()
    
    NPAs = ResultData("NumPAs")
    ImpPAsSAP = ResultData("ImpPAsSAP")
    ImpDifSAP = ResultData("ImpDifSAP")
    ImpUSER = ResultData("ImpUSER")
    aux = False
    If ImpDifSAP <> 0 Then
        Invoinces = FormatNumber(Invoinces, 2)
        DifInvoices = ImpPAsSAP - Invoinces
        DifInvoices = FormatNumber(DifInvoices, 2)
        NInvoices = FormatNumber(dicInvoices.Count, 2)
        If DifInvoices <> 0 And NPAs <> NInvoices Then
            ask = MsgBox("Hay diferencia en las facturas. Total: " & DifInvoices & vbCrLf & _
                        "¿Quiere ajustar la diferencia?", vbYesNo, "Confirmacion")
            If ask = vbNo Then Exit Sub
            Dim PAs As Object
            Set PAs = CreateObject("Scripting.Dictionary")
            Set PAs = ItemsFoundSAP(Ses)
            
            For Each inv In dicInvoices.keys
                If Not PAs.exists(inv) Then
                    InvVal = application.WorksheetFunction.VLookup(inv, PDws.Range("B1:D" & EndRow & ""), 3, False)
                    InvSodd = UCase(WorksheetFunction.VLookup(inv, PDws.Range("B1:I" & EndRow & ""), 8, False))
                    For Each ClientName In Clients.keys
                        If InStr(1, InvSodd, ClientName) > 0 Then
                            InvCode = Clients(ClientName)
                            If InvVal < 0 Then
                                InvVal = InvVal * -1
                                InvComent = "SE DESCUENTA ABONO " & inv
                                ModuleSAP.NewEntry "06", InvCode
                                ModuleSAP.NewEntryAddData InvVal, VTO, InvComent, -1, AsignationVTO
                            ElseIf InvVal > 0 Then
                                InvComent = "PAGA FACTURA " & inv
                                ModuleSAP.NewEntry "16", InvCode
                                ModuleSAP.NewEntryAddData InvVal, VTO, InvComent, -1, AsignationVTO
                            End If
                        End If
                    Next ClientName
                End If
            Next
        ElseIf DifInvoices <> 0 And NPAs = NInvoices Then
            MsgBox "La diferencia esta en centimos", , "Diferencia"
            DifConfirmation.DifMSG DifInvoices
        End If
    End If
    Positions = ModuleSAP.Simulate()
    For i = Positions(0) + 1 To Positions(1)
        ModuleSAP.EnterPosition i
        ModuleSAP.NewEntryAddData 0, 0, ChequeComentary, -1, AsignationVTO
    Next
    
    ModuleAux.SaveConfirmation
    ModuleSAP.CallTransaction "FB03"
    EntryNumber = ModuleSAP.GetEntryNumber
    Path = PDwb.Path
    PDwb.SaveAs Filename:=Path & "\" & EntryNumber & " ClientName " & UserAmount & ".xlsx", FileFormat:=51
    MsgBox ("Se ha aplicado el pago con nº asiento " & EntryNumber & " Y se ha guardado el archivo")
    PDwb.Close SaveChanges:=False
    Kill DetailPath
    Exit Sub
    
ErrorHandler:
    If Err.Number = 6 Then
        Resume Next
        Err.Clear
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If


End Sub
