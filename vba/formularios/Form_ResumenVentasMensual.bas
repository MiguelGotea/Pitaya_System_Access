' ==========================================================
' Modulo  : Form_ResumenVentasMensual
' Tipo    : 100
' Lineas  : 84
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:23
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
Else
    mesac.Value = mesac.Value - 1
End If

End Sub

Private Sub Comando142_Click()
If mesac.Value = 12 Then
    mesac.Value = 1
    añoac.Value = añoac.Value + 1
Else
    mesac.Value = mesac.Value + 1
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
archivo = Direccionmes & "\" & Me.mesac & "_" & Me.añoac & " - " & codigoLocal() & " - Ventas Acumuladas Totales.pdf"
DoCmd.OutputTo acOutputForm, "ResumenVentasMensual", acFormatPDF, archivo, True

End Sub

Private Sub Comando65_Click()
Dim sobrantefaltante As Double
Dim ventatarjeta, ventaefectivo, ventahugo, ventapedidosya, ventaotro As Double
Me.Requery

Me.ingresoxventas = VentasTotalesMesGuardado(Me.mesac, Me.añoac)
Me.CantidadTipo.ControlSource = "=FRVentaxGrupoMesLocal([mesac],[añoac],[CodGrupo],0)"
Me.MontoTipo.ControlSource = "=FRVentaxGrupoMesLocal([mesac],[añoac],[CodGrupo],3)"


''''''''''''''''''''''''''''''''''''''

Me.cierres = SumaCierresMesInterno(Me.mesac, Me.añoac)
Me.cajainicial = SumaCajaInicialMesInterno(Me.mesac, Me.añoac)
Me.compras = SalidasDeCajaMesInterno(Me.mesac, Me.añoac)

ventatarjeta = FRVentaInternoMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, -1)
ventaefectivo = FRVentaInternoMesTipoDelivGuardado(Me.mesac, Me.añoac, 0, 0)
ventahugo = FRVentaInternoMesTipoDelivGuardado(Me.mesac, Me.añoac, 5, -1)
ventapedidosya = FRVentaInternoMesTipoDelivGuardado(Me.mesac, Me.añoac, 8, -1)
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


