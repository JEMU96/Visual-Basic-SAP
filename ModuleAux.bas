Attribute VB_Name = "ModuleAux"
Function OpenFile(Msg As String)
    Do While OpenFile = False
        OpenFile = application.GetOpenFilename(Title:=Msg)
        If OpenFile = False Then
            ask = MessageBox(&H0, "No se han seleccionado fichero." & vbCrLf & _
        "¿Quiere continuar?", "Confirmación", vbRetryCancel)
            If ask = vbCancel Then
                MsgBox "Proceso cancelado por el usuario"
                Exit Do
                Exit Function
            End If
        End If
    Loop
End Function
Function AskUserDate()
    ask = False
    Do While Not IsDate(ask)
        ask = InputBox("Introduce la fecha en formato dd/mm/yyyy.", "Introduce Fecha")
        If IsDate(ask) Then
                AskUserDate = ask
                Exit Function
        Else
            MsgBox "Fecha no Valida."
        End If
    Loop
End Function
Function AskUserNumber(Msg)
    ask = False
    Do While ask = False
        ask = application.InputBox("Introduce el importe de " & Msg & ".", "Introduce importe", Type:=1)
        If IsNumeric(ask) Then
            On Error Resume Next
            ask = Replace(ask, ".", ",")
            AskUserNumber = FormatNumber(ask, 2)
            Exit Function
        Else
            MsgBox "Importe no valido."
        End If
    Loop
End Function
Function AskUserString(Msg)
    ask = False
    Do While ask = False
        ask = application.InputBox("Introduce el " & Msg, "Introduce Comentario", Type:=2)
        If ask <> "" Then
                AskUserString = ask
                Exit Function
        Else
            MsgBox "No valido."
            ask = False
        End If
    Loop
End Function
Function SaveConfirmation()
    ask = MessageBox(&H0, "Conprueba los apuntes." & vbCrLf & _
    "Si son correctos guarde en SAP y Acepte.", "Confirmación", vbYesNo)
    If ask = vbNo Then
        MsgBox "Proceso cancelado por el usuario", vbOKOnly, "Cancelación"
        BackToMain
        Continue = False
        Exit Function
    ElseIf ask = vbYes Then
        titl = ChkWindow()
        Do While InStr(1, titl, "Visualizar Resumen") > 0
            ask = MessageBox(&H0, "Conprueba los apuntes." & vbCrLf & _
                "Si son correctos guarde en SAP y Acepte.", "Confirmación", vbYesNo)
            If ask = vbCancel Then
                MsgBox "Proceso cancelado por el usuario", vbOKOnly, "Cancelación"
                BackToMain
                Continue = False
                Exit Function
            Else
                titl = ChkWindow()
             End If
        Loop
        'ses.findById("wnd[0]/tbar[0]/btn[11]").press ' Auto-Save
    End If
End Function

