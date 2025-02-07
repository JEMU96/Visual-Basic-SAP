Attribute VB_Name = "ModuleConf"
Sub ConfirminOneClient()
    On Error GoTo ErrorHandler
    Continue = True
    'Client Variables
    Client = "12345" ' Sensitive data
    ClientName = "ClientName" 'Sensitive data
    'Public Variable
    MainAccount = Client
    Dim dicInvoices As Object
    Set dicInvoices = CreateObject("Scripting.Dictionary")
    
    Dim PDwb As Workbook
    Dim PDws As Worksheet
    DetailPath = ModuleAux.OpenFile("Abre la realcion de pago")
    If DetailPath = False Then Exit Sub
    Set PDwb = Workbooks.Open(DetailPath)
    Set PDws = PDwb.Sheets(1)
    EndRow = PDws.Cells(Rows.Count, 1).End(xlUp).Row
    IniRow = 2
    
    Dim Template As String
    Template = "C:\servername\Template.xlsx" 'Auto select the template from the Business Server
    'Tmeplate = ModuleAux.OpenFile("Abre la Plantilla Call Transaction") ' Ask User to open Template
    Dim Tempwb As Workbook
    Dim Tempws As Worksheet
    Set Tempwb = Workbooks.Open(Template)
    Set Tempws = Tempwb.Sheets(1)
    
    
    UserAmount = ModuleAux.AskUserNumber("Introduce el total del Pagaré")
    PDAmount = FormatNumber(application.Sum(PDws.Columns("O:O")), 2)
    
    If UserAmount <> PDAmount Then
        MsgBox "El importe no cuadra. Se cancela el proceso", , "Cancelación"
        Exit Sub
    End If
    
    TempIniRow = 10
    TempEndRow = Tempws.Cells(Rows.Count, 4).End(xlUp).Row
    If TempIniRow < TempEndRow Then
        Tempws.Range("D10:D" & TempEndRow & "").Clear
    End If
    PDws.Range("L:L").Replace what:="-", Replacement:="", LookAt:=xlPart
        
    VTO = PDws.Cells(EndRow, 1).Value
    AsignationVTO = Format(VTO, "yyyymmdd")
    VTO = Format(VTO, "dd.mm.yyyy")
    
    DocDate = Date
    AsignationDoc = Format(DocDate, "yyyymmdd")
    DocDate = Format(DocDate, "dd.mm.yyyy")
    NPayment = PDws.Cells(2, 1).Value
    
    Invoices = FormatNumber(0, 2)
    Cargos = FormatNumber(0, 2)
    Abonos = FormatNumber(0, 2)
    
    ChequeComentary = "PAG. " & ClientName & " " & NPayment & " VTO. " & VTO
    CargosComentary = "TOTAL CARGOS " & ClientName & " " & NPayment & " VTO. " & VTO
    AbonosComentary = "TOTAL ABONOS " & ClientName & " " & NPayment & " VTO. " & VTO
    AJDComentary = "GASTOS AJD " & ClientName & " " & NPayment & " VTO. " & VTO

    For i = 2 To EndRow
        If InStr(1, PDws.Cells(i, 11).Value, "FACTURA") > 0 And (InStr(1, PDws.Cells(i, 12).Value, "V") Or InStr(1, PDws.Cells(i, 12).Value, "X") Or InStr(1, PDws.Cells(i, 12).Value, "Y")) > 0 Then
            Tempws.Cells(TempIniRow, 4).Value = PDws.Cells(i, 12).Value
            dicInvoices.Add PDws.Cells(i, 12).Value, PDws.Cells(i, 15).Value
            TempIniRow = TempIniRow + 1
            Invoices = Invoices + PDws.Cells(i, 15).Value
        ElseIf InStr(1, PDws.Cells(i, 11).Value, "FACTURA") > 0 And (InStr(1, PDws.Cells(i, 12).Value, "V") And InStr(1, PDws.Cells(i, 12).Value, "X") And InStr(1, PDws.Cells(i, 12).Value, "Y")) = 0 Then
            Cargos = Cargos + PDws.Cells(i, 15).Value
        ElseIf InStr(1, PDws.Cells(i, 11).Value, "CARGO") > 0 Then
            Cargos = Cargos + PDws.Cells(i, 15).Value
        ElseIf InStr(1, PDws.Cells(i, 11).Value, "ABONO") > 0 Then
            Abonos = Abonos + PDws.Cells(i, 15).Value
        End If
    Next i
    Tempws.Cells(2, 5).Value = DocDate
    Tempws.Cells(2, 7).Value = DocDate
    Tempwb.Save
    Tempwb.Close
    
    Set Ses = ModuleSAP.ConnectToSAP ' Connect to SAP GUI
    ModuleSAP.BatchInput Template
    ModuleSAP.NewEntry "90", MainAccount, "K", DocDate '90 K Business Confirmin Entry Codification "Sensitive Data"
    ModuleSAP.NewEntryAddData UserAmount, VTO, ChequeComentary, -1, AsignationDoc
    
    If Abonos <> 0 Then
        ModuleSAP.NewEntry "61", Client '61 Business Debit Entry Codification "Sensitive Data"
        ModuleSAP.NewEntryAddData Abonos, VTO, AbonosComentary, -1, AsignationVTO
    End If
    If Cargos <> 0 Then
        Cargos = Cargos * -1
        ModuleSAP.NewEntry "60", Client '60 Business Credit Entry Codification "Sensitive Data"
        ModuleSAP.NewEntryAddData Cargos, VTO, CargosComentary, -1, AsignationVTO
    End If
    
    Set ResultData = SAPData()
    
    NPAs = ResultData("NumPAs")
    ImpPAsSAP = ResultData("ImpPAsSAP")
    ImpDifSAP = ResultData("ImpDifSAP")
    ImpUSER = ResultData("ImpUSER")
    If ImpDifSAP <> 0 Then
        Invoices = FormatNumber(Invoices, 2)
        DifInvoices = ImpPAsSAP - Invoices
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
                    InvVal = dicInvoices(inv)
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
            Next inv
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
    PDwb.SaveAs Filename:=Path & "\" & EntryNumber & ClientName & UserAmount & ".xlsx", FileFormat:=51
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

