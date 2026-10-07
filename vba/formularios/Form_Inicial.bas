' ==========================================================
' Modulo  : Form_Inicial
' Tipo    : 100  |  Lineas: 52
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database

Private Sub Comando24_Click()
DoCmd.Save
Me.Requery
If cajainicial(Date) = 0 Then
    MsgBox "Ingresar Caja Inicial"
    Call Comando25_Click
    [Forms]![Inicial].Form.SetFocus
    [Forms]![Contador Caja].Form.SetFocus
Else
    If MontoCierre(Me.Fecha - 1, "T") <> Me.Dinero Then
        MsgBox "El dinero ingresado no coincide con el monto dejado en el cierre anterior, verifficar conteo nuevamente"
        DoCmd.OpenForm "Contador Caja", acNormal, , , acFormAdd
        [Forms]![Contador Caja]![titulo1] = "CONTEO DE CAJA INICIAL"
    Else
        'cuando ya cargo correctamente la caja inicial
        'dESCARGAR csv DE DIA ANTERIOR
        Call SincronizarDatosClientesLocales
        Call ActualizarArchivoCSVVentasFecha(Date)
        Call ActualizarArchivoCSVVentasFecha(Date - 1)
        Call ActualizarArchivoCSVMembresias
        
        DoCmd.Close acForm, "Inicial"
        'Application.FollowHyperlink "https://erp.batidospitaya.com"
    End If
End If

End Sub

Public Sub Comando25_Click()
If IsNull(Me.Dinero) Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO Estadoinicial(Dinero, Fecha, [TipoCambio$_C$]) values " & _
    "(0, #" & Date & "#, " & tipocambio(Date) & ")"
    DoCmd.SetWarnings True

    Me.Requery
    DoCmd.OpenForm "Contador Caja", acNormal, , , acFormAdd
    [Forms]![Contador Caja]![titulo1] = "CONTEO DE CAJA INICIAL"
ElseIf Me.Dinero = 0 Then
    DoCmd.OpenForm "Contador Caja", acNormal, , , acFormAdd
    [Forms]![Contador Caja]![titulo1] = "CONTEO DE CAJA INICIAL"
Else
    MsgBox "Ya existe valor ingresado de Caja Inicial, no se puede cambiar"
    
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
