Attribute VB_Name = "ModuleDiario"
Sub BankFile()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wb_Yest As Workbook
    Dim ws_Yest As Worksheet
    BankFilePath = ModuleAux.OpenFile("Abre el fichero del banco de hoy")
    Set wb = Workbooks.Open(BankFilePath)
    Set ws = wb.Sheets(1)
    
    
    Rows("1:7").EntireRow.Delete ' Bank Head  can erease that data
    ws.Columns("C:C").Replace what:=".", Replacement:="", LookAt:=xlPart ' Description Column errease dots to shorten lenght
    YestDailyPath = ModuleAux.OpenFile("Abre los pagos del último día")
    wb_YestName = Workbooks.Open(YestDailyPath).Name
    Set wb_Yest = Workbooks(wb_YestName)
    Set ws_Yest = wb_Yest.Sheets(1)
    LastPay = ws_Yest.Cells(2, 2).Value
    Dim DicNoAply As Object
    Set DicNoAply = CreateObject("Scripting.Dictionary")
    DicNoAply.RemoveAll
    Set Dato = ws.Range("C:C").Find(LastPay, LookIn:=xlValues, LookAt:=xlWhole)
    EndRowYest = Dato.Row
    For i = 2 To EndRowYest
        If ws_Yest.Cells(i, 12).Text = "No Aplicado" Then
            DicNoAply.Add i, i
        End If
    Next
    ws.Activate
    EndRow = ws.Rows.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Rows(EndRowYest & ":" & EndRow).EntireRow.Delete
    'Delete extra columns
    ws.Columns("B:B").Delete
    ws.Columns("D:F").Delete
    ws.Columns("E:G").Delete
    
    Dim Celda As Range
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ValList = Array("SOLO", "TODO", "HASTA", "ENTRE", "RELACION", "REEMBOLSO", "A CUENTA", "FACTURA") 'Validation List to let the user choose what to do with each payment
    ValList = Join(ValList, ",")
    With ws.Range("I2:I" & EndRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=ValList
    End With
    For i = EndRow To 2 Step -1
        If Cells(i, 3).Value < 0 Then
            ws.Rows(i).Delete
        Else
            Dim CellValue As String
            CellValue = ws.Cells(i, 2).Value
            If InStr(1, CellValue, "Ingreso", vbTextCompare) > 0 Then
                ws.Cells(i, 6).Value = CellValue
                ws.Cells(i, 7).FormulaR1C1 = "=Len(RC[-1])"
                ws.Cells(i, 8).Value = Right(ws.Cells(i, 1), 4) & Mid(ws.Cells(i, 1), 4, 2) & Left(ws.Cells(i, 1), 2)
            ElseIf InStr(1, Cells(i, 2).Value, "Transferencia De", vbTextCompare) > 0 Then
                Dim startIndex As Long
                Dim endIndex As Long
                startIndex = InStr(CellValue, "De") + 3
                endIndex = InStr(startIndex, CellValue, ",") - 1
                ws.Cells(i, 6).Value = "Transf " & Mid(CellValue, startIndex, endIndex - startIndex + 1) & " " & ws.Cells(i, 1).Value
                ws.Cells(i, 7).FormulaR1C1 = "=Len(RC[-1])"
                ws.Cells(i, 8).Value = Right(ws.Cells(i, 1), 4) & Mid(ws.Cells(i, 1), 4, 2) & Left(ws.Cells(i, 1), 2)
            Else
                ws.Rows(i).Delete
            End If
        End If
    Next i
    'Head Columns
    ws.Cells(1, 6).Value = "Concepto"
    ws.Cells(1, 7).Value = "Largo"
    ws.Cells(1, 8).Value = "Asignacion"
    ws.Cells(1, 9).Value = "Accion"
    ws.Cells(1, 10).Value = "Vto. Final"
    ws.Cells(1, 11).Value = "Vto. Inicial"
    ws.Cells(1, 12).Value = "¿APLICADO?"
    ws.Cells(1, 13).Value = "Nº ASIENTO"
    EndRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim Rango As Range
    Set Rango = Range("G2:G" & EndRow1)
    Rango.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="=50"
    Rango.FormatConditions(Rango.FormatConditions.Count).SetFirstPriority
    With Rango.FormatConditions(1).Font
        .Color = -11489280 ' Green Color
    End With
    With Rango.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798 ' Green Color
        .TintAndShade = 0
    End With
    ' Format if > 50 (Red)
    Rango.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=50"
    Rango.FormatConditions(Rango.FormatConditions.Count).SetFirstPriority
    With Rango.FormatConditions(1).Font
        .Color = -16776961 ' Red Color
    End With
    With Rango.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615 ' Red Color
        .TintAndShade = 0
    End With
    Dim aux() As Long
    ReDim aux(1 To ActiveSheet.UsedRange.Rows.Count)
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        If ws.Range("G" & i).Value > 50 Then
            aux(i) = i
        End If
    Next i
    If DicNoAply.Count > 0 Then
        Matriz = DicNoAply.keys
        For i = 0 To DicNoAply.Count - 1
            YestRow = Matriz(i)
            ws_Yest.Rows(YestRow).EntireRow.Copy
            EndRowNow = ws.UsedRange.Rows.Count + 1
            ws.Rows(EndRowNow).PasteSpecial Paste:=xlPasteAll
        Next i
    End If
    wb_Yest.Close
    wb.Save
    wb.Close
    MsgBox "Selecciona que hacer con cada pago", , "Fin"
End Sub

Sub DailyPayments()
    Dim ws As Worksheet
    Dim wb As Workbook
    BankFilePath = ModuleAux.OpenFile("Abre el fichero del banco de hoy Tratado")
    Set wb = Workbooks.Open(BankFilePath)
    Set ws = wb.Sheets(1)
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    BankAccount = "5555555" ' SAP Bank Account "Sensitive Data"
    AccConfirmation = MsgBox("¿Hay alguna cuenta de Acreedor?", vbYesNo, "Confirmacion")
    ClientCategory = "D"
    Set Ses = ModuleSAP.ConnectToSAP ' Connet to SAP GUI
    For i = EndRow To 2 Step -1
        If ws.Range("l" & i).Text = "Aplicado" Then
            MsgBox ("Pago ya aplicado, se pasa al siguiente")
            GoTo EndLine
        End If
        'Variables
        Continue = True
        DocDate = ws.Range("A" & i).Value
        VTO = Format(DocDate, "dd.mm.yyyy")
        Asignation = Format(DocDate, "yyyymmdd")
        Amount = FormatNumber(ws.Cells(i, 3).Value, 2)
        Comentary = ws.Range("F" & i).Value
        ClientCode = ws.Cells(i, 4).Value
        MainAccount = ClientCode
        SearchData1 = ws.Range("j" & i).Value
        SearchData2 = ws.Range("k" & i).Value
        
        If AccConfirmation = vbYes Then
            ClientCategory = InputBox("Introduce la Categoria del Cliente(Acreedor K, Deudor D), Cliente Nº " & ws.Cells(i, 4))
        End If
        
        If ws.Range("i" & i) = "" Then
            MsgBox ("No hay partidas indicadas en fila " & i)
            ws.Cells(i, 12).Value = "No Aplicado"
            GoTo EndLine
        ElseIf UCase(ws.Range("i" & i)) = "RELACION" Then
            PaymentDetailPath = ModuleAux.OpenFile("Abre el archivo con el detalle de Facturas " & ws.Range("f" & i) & " " & ws.Range("c" & i))
            Set wb_PaymentDetailPath = Workbooks.Open(PaymentDetailPath)
            Template = ModuleAux.OpenFile("Abre la Plantilla de Call Transaction")
            Set wb_Template = Workbooks.Open(Template)
            TempIniRow = 10
            TempEndRow = wb_Template.Sheets(1).Cells(Rows.Count, "D").End(xlUp).Row
            If TempIniRow < TempEndRow Then
                wb_Template.Sheets(1).Range("D" & TempIniRow & ":D" & TempEndRow).Delete
            End If
            wb_Template.Sheets(1).Cells(2, 5).Value = DocDate
            wb_Template.Sheets(1).Cells(2, 7).Value = DocDate
            wb_Template.Sheets(1).Cells(6, 6).Value = ClientCategory
            application.InputBox(prompt:="Selecciona las facturas que se pasaran en la Template", Type:=8).Copy
            wb_Template.Sheets(1).Range("D" & TempIniRow).PasteSpecial xlPasteValues
            wb_Template.Save
            wb_Template.Close
            ModuleSAP.BackToMain Template
            ModuleSAP.NewEntry "40", BankAccount, , VTO
            ModuleSAP.NewEntryAddData Amount, VTO, Comentary, -1, Asignation
            aux = vbYes
            Do While aux = vbYes
                ask = MsgBox("¿Tiene algún apunte manual?", vbYesNo)
                If ask = vbYes Then
                    VAC = ModuleAux.AskUserNumber("Introduce el Importe del Apunte: ")
                    If VAC < 0 Then
                        VAC = VAC * -1
                        ModuleSAP.NewEntry "60", Client
                        ModuleSAP.NewEntryAddData VAC, VTO, Comentary, -1, Asgination
                    ElseIf VAC > 0 Then
                        ModuleSAP.NewEntry "61", Client
                        ModuleSAP.NewEntryAddData VAC, VTO, Comentary, -1, Asgination
                    End If
                    aux = MsgBox("¿Hay mas apuntes manuales?", vbYesNo)
                Else
                    Exit Do
                End If
                On Error Resume Next
            Loop
            wb_PaymentDetailPath.Close
            Positions = ModuleSAP.Simulate()
            For i = Positions(0) + 1 To Positions(1)
                ModuleSAP.EnterPosition i
                ModuleSAP.NewEntryAddData 0, 0, ChequeComentary, -1, AsignationVTO
            Next
            ModuleAux.SaveConfirmation
            EntryNum = ModuleSAP.GetEntryNumber()
            ModuleSAP.SaveEntry
            ws.Cells(i, 12).Value = "Aplicado"
            ws.Cells(i, 13).Value = EntryNum
            GoTo EndLine
        Else
            ModuleSAP.CallTransaction "F-04"
            ModuleSAP.NewEntry "40", BankAcount, , VTO
            If UCase(ws.Range("i" & i)) = "FACTURA" Then
                ModuleSAP.SearchItems ClientCategory, 5, SearchData1, , ClientCode ' 5 Invoice selection
                If Continue = False Then GoTo EndLine
            ElseIf UCase(ws.Range("i" & i)) = "TODO" Then
                ModuleSAP.SearchItems ClientCategory, 0, , , ClientCode ' 0 All items
                If Continue = False Then GoTo EndLine
            ElseIf UCase(ws.Range("i" & i)) = "HASTA" Then
                ModuleSAP.SearchItems ClientCategory, 16, , , ClientCode, SearchData1 ' 16 VTOs
                If Continue = False Then GoTo EndLine
            ElseIf UCase(ws.Range("i" & i)) = "SOLO" Then
                  ModuleSAP.SearchItems ClientCategory, 16, SearchData1, , ClientCode ' 16 VTOs
                If Continue = False Then GoTo EndLine
            ElseIf UCase(ws.Range("i" & i)) = "ENTRE" Then
                  ModuleSAP.SearchItems ClientCategory, 16, SearchData1, , ClientCode, SearchData2 ' 16 VTOs
                If Continue = False Then GoTo EndLine
            ElseIf UCase(ws.Range("i" & i)) = "A CUENTA" Then
                  ModuleSAP.NewEntry "16", ClientCode, , VTO
                  ModuleSAP.NewEntryAddData Amount, VTO, Comentary, , AsignationVTO
                If Continue = False Then GoTo EndLine
            ElseIf UCase(ws.Range("i" & i)) = "REEMBOLSO" Then
                Today = Date
                MonthToday = Month(Today)
                If Left(Today, 2) < 8 Then
                    FechaReem = "25." & Month(Today) & "." & Year(Today)
                Else
                    If MonthToday = 12 Then
                        FechaReem = "25.01." & Year(Today) + 1
                    Else
                        FechaReem = "25." & Month(Today) + 1 & "." & Year(Today)
                    End If
                End If
                AsignationReem = Right(FechaReem, 4) & Mid(FechaReem, 4, 2) & Left(FechaReem, 2)
                ComentaryReem = "Tran. Reemb. Dronas OS " & ws.Range("j" & i).Value
                ModuleSAP.NewEntry "36", ClientCode, , VTO
                ModuleSAP.NewEntryAddData Amount, FechaReem, Comentary, -1, AsignationReem
                If Continue = False Then GoTo EndLine
            End If
            Set ResultData = SAPData()
            NPAs = ResultData("NumPAs")
            ImpPAsSAP = ResultData("ImpPAsSAP")
            ImpDifSAP = ResultData("ImpDifSAP")
            ImpUSER = ResultData("ImpUSER")
        End If
    
EndLine:
    Next i
    MsgBox ("Pagos Diarios Aplicados: Comprueba los nº de Asientos")
    ws.Columns("M:M").AutoFit
    End Sub

