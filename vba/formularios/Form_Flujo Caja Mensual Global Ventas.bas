' ==========================================================
' Modulo  : Form_Flujo Caja Mensual Global Ventas
' Tipo    : 100  |  Lineas: 98
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database
Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 10500
End Sub
Private Sub Comando139_Click()
If mesac.Value = 1 Then
    mesac.Value = 12
    añoac.Value = añoac.Value - 1
    Me.ingresosactual = 0
Else
    mesac.Value = mesac.Value - 1
    Me.ingresosactual = 0
End If

End Sub

Private Sub Comando142_Click()
If mesac.Value = 12 Then
    mesac.Value = 1
    añoac.Value = añoac.Value + 1
    Me.ingresosactual = 0
Else
    mesac.Value = mesac.Value + 1
    Me.ingresosactual = 0
End If
End Sub



Private Sub Comando574_Click()
Call Comando65_Click

Dim Direccionano, Direccionmes As String
Dim archivo As String
Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.añoac
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.añoac & "\" & Me.mesac
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If
archivo = Direccionmes & "\" & Me.mesac & "_" & Me.añoac & " - Ventas Acumuladas Totales.pdf"
DoCmd.OutputTo acOutputForm, "Flujo Caja Mensual Global Ventas", acFormatPDF, archivo, True

End Sub

Private Sub Comando65_Click()
Dim sobrantefaltante As Double
Dim ventatarjeta, ventaefectivo, ventahugo, ventapedidosya, ventaotro As Double
Me.Requery

Me.VentasTotal2 = VentasTotalesMesLocalGuardado(Me.mesac, Me.añoac, 2)
Me.VentasTotal4 = VentasTotalesMesLocalGuardado(Me.mesac, Me.añoac, 4)
Me.VentasTotal5 = VentasTotalesMesLocalGuardado(Me.mesac, Me.añoac, 5)
Me.ingresoxventas = Me.VentasTotal2 + Me.VentasTotal4 + Me.VentasTotal5

Me.CantidadTipo2.ControlSource = "=FRVentaxGrupoMes([mesac],[añoac],[CodGrupo],0,2)"
Me.CantidadTipo4.ControlSource = "=FRVentaxGrupoMes([mesac],[añoac],[CodGrupo],0,4)"
Me.CantidadTipo5.ControlSource = "=FRVentaxGrupoMes([mesac],[añoac],[CodGrupo],0,5)"

Me.MontoTipo2.ControlSource = "=FRVentaxGrupoMes([mesac],[añoac],[CodGrupo],3,2)"
Me.MontoTipo4.ControlSource = "=FRVentaxGrupoMes([mesac],[añoac],[CodGrupo],3,4)"
Me.MontoTipo5.ControlSource = "=FRVentaxGrupoMes([mesac],[añoac],[CodGrupo],3,5)"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.cierres = SumaCierresMesGlobal(Me.mesac, Me.añoac)
Me.cajainicial = SumaCajaInicialMes(Me.mesac, Me.añoac)
Me.compras = SalidasDeCajaMes(Me.mesac, Me.añoac)

ventatarjeta = FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, -1, 2) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, -1, 4) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, -1, 5)
ventaefectivo = FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, 0, 2) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, 0, 4) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, 0, 5)
ventahugo = FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 5, -1, 2) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 5, -1, 4) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 5, -1, 5)
ventapedidosya = FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 8, -1, 2) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 8, -1, 4) + FRVentaTotalMesTipoDelivGuardado(Me.mesac, Me.añoac, 8, -1, 5)
ventaotro = Me.ingresoxventas - ventatarjeta - ventaefectivo - ventahugo - ventapedidosya

Me.ventaspos = ventatarjeta 'totalxventa2pos + totalxventa4pos 'puede ser FRVentasTotalGruposMes(Me.mesac, Me.añoac, 1)
Me.ventashugo = ventahugo
Me.ventaspedidosya = ventapedidosya
Me.ventasotro = ventaotro
Me.ingresosactual = Me.cierres - Me.cajainicial + Me.compras + Me.ventaspos + Me.ventashugo + Me.ventaspedidosya + Me.ventasotro

sobrantefaltante = Me.ingresosactual - Me.ingresoxventas
If sobrantefaltante > 0 Then
    Me.sobrafaltaetiqueta.Caption = "SOBRANTE"
    Me.sobrafalta = sobrantefaltante
Else
    Me.sobrafaltaetiqueta.Caption = "FALTANTE"
    Me.sobrafalta = (-1) * sobrantefaltante
End If


End Sub


