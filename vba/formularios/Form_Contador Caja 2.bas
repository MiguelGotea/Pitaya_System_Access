' ==========================================================
' Modulo  : Form_Contador Caja 2
' Tipo    : 100  |  Lineas: 377
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database


Private Function ResumenTotalesComprasDia(fechi As Date) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT compras.CodIngresoAlmacen, compras.Fecha, compras.Tipo, compras.Cantidad, compras.CostoTotal," & _
" nombreproductocoti([Compras]![CodCotizacion]) AS [names]" & _
" FROM compras WHERE (((compras.Fecha) = #" & fechi & "#) And ((compras.Tipo) = 'CAJA'))" & _
" ORDER BY compras.CodIngresoAlmacen"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
ResumenTotalesComprasDia = ""

Do While Not rst.EOF
    ResumenTotalesComprasDia = ResumenTotalesComprasDia & rst("Cantidad") & vbTab & rst("names") & " = " & rst("CostoTotal") & vbCrLf
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
ResumenTotalesComprasDia = ""
End Function
Private Function ResumenProductosComprasDia(fechi As Date) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT compras.CodIngresoAlmacen, compras.Fecha, compras.Tipo," & _
" nombreproductocoti([Compras]![CodCotizacion]) AS [names]" & _
" FROM compras WHERE (((compras.Fecha) = #" & fechi & "#) And ((compras.Tipo) = 'CAJA'))" & _
" ORDER BY compras.CodIngresoAlmacen"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
ResumenProductosComprasDia = ""

Do While Not rst.EOF
    ResumenProductosComprasDia = ResumenProductosComprasDia & rst("names") & vbCrLf
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
ResumenProductosComprasDia = ""
End Function

Private Function ResumenCantidadComprasDia(fechi As Date) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT compras.CodIngresoAlmacen, compras.Fecha, compras.Tipo, compras.Cantidad" & _
" FROM compras WHERE (((compras.Fecha) = #" & fechi & "#) And ((compras.Tipo) = 'CAJA'))" & _
" ORDER BY compras.CodIngresoAlmacen"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
ResumenCantidadComprasDia = ""

Do While Not rst.EOF
    ResumenCantidadComprasDia = ResumenCantidadComprasDia & rst("Cantidad") & vbCrLf
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
ResumenCantidadComprasDia = ""
End Function

Private Function ResumenMontoComprasDia(fechi As Date) As String
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT compras.CodIngresoAlmacen, compras.Fecha, compras.Tipo, compras.CostoTotal" & _
" FROM compras WHERE (((compras.Fecha) = #" & fechi & "#) And ((compras.Tipo) = 'CAJA'))" & _
" ORDER BY compras.CodIngresoAlmacen"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

contador = 0
ResumenMontoComprasDia = ""

Do While Not rst.EOF
    ResumenMontoComprasDia = ResumenMontoComprasDia & rst("CostoTotal") & vbCrLf
    contador = contador + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
ResumenMontoComprasDia = ""
End Function









Private Sub Comando1578_Click()
DoCmd.OpenForm "HistorialVentasPedido"
[Forms]![HistorialVentasPedido].horalimite = Me.ahoracierre
[Forms]![HistorialVentasPedido].horadesde = Me.ahoracierreinicio
[Forms]![HistorialVentasPedido].fechasistema.Enabled = False

[Forms]![HistorialVentasPedido].DetalleHoras.width = 0
[Forms]![HistorialVentasPedido].tieneeliminados.width = 0
[Forms]![HistorialVentasPedido].Caja.width = 0
[Forms]![HistorialVentasPedido].HorasAnulacion.width = 0
End Sub

Private Sub Comando2599_Click()
DoCmd.OpenForm "Ingreso de Compras"
[Forms]![Ingreso de Compras]![operarioorigen] = [Forms]![Main Pitaya]![codigologin]


End Sub

Private Sub Comando823_Click()
'Alertas
If Me.tventasposs = 0 Then
    If MsgBox("No se ingreso ningun valor en ventas con POS, CONFIRMA que no se realizo ninguna venta con tarjeta por el POS?", vbYesNo, "Confirmar ventas cero") = vbNo Then
        Exit Sub
    End If
End If
If Me.tventaspedidosya = 0 Then
    If MsgBox("No se ingreso ningun valor en ventas con PedidosYa, CONFIRMA que no ingreso ningun pedido de PedidosYa?", vbYesNo, "Confirmar ventas cero") = vbNo Then
        Exit Sub
    End If
End If
If Me.tventaspedidosya = 0 Then
    If MsgBox("No se ingreso ningun valor en ventas con Transferencia, CONFIRMA que no se realizo ninguna venta pagada por trasnferencia?", vbYesNo, "Confirmar ventas cero") = vbNo Then
        Exit Sub
    End If
End If


'Datos basicos de cierre
Dim dcierre As Integer
Dim dinicio As Date
Dim dfinal As Date
Dim dfecha As Date
Dim dopera As Integer

Dim gventaspos As Double
Dim gventastrans As Double
Dim gventaspya As Double
Dim gcordoba As Double
Dim gdola As Double
Dim pagoscaja As Double
Dim aligera As Double
Dim inici As Double
Dim gsobrafalta As Double
Dim tipocam As Double

gventaspos = Me.tventasposs
gventastrans = Me.tventastransferencia
gventaspya = Me.tventaspedidosya
gcordoba = Me.scor
gdola = Me.sdol

dcierre = Me.codigocierrex
dinicio = Me.ahoracierreinicio
dfinal = Me.ahoracierre
dfecha = Date
dopera = Me.acodigocajero
pagoscaja = pagospordia(Date)
aligera = TotalRetirosCajaDiaEquivalente(Date)
inici = cajainicial(Date)
tipocam = tipocambio(Date)

' actualizar ventana de cierre de caja 2
Dim pedidosguardados As Double
Dim pedidosguardadospos As Double
Dim pedidosguardadostrasnfer As Double
Dim pedidosguardadospedidosya As Double
Dim pedidosguardadosefectivo As Double
'Dim comprascaja As Double

pedidosguardados = AcumuladoDiaHastaHoraGuardado(Date, dfinal) 'AcumuladoDiaGuardado(Date)
pedidosguardadospos = montosoloposdiahastahoraguardado(Date, dfinal) 'montosoloposdiaguardado(Date)
pedidosguardadostrasnfer = montoposdiahastahoraguardadoTransferencia(Date, dfinal) 'montoposdiaguardadoTransferencia(Date)
pedidosguardadospedidosya = montohugodiahastahoraguardado(Date, dfinal) 'montohugodiaguardado(Date)
pedidosguardadosefectivo = pedidosguardados - pedidosguardadospos - pedidosguardadostrasnfer - pedidosguardadospedidosya
gsobrafalta = gcordoba + gdola * tipocam - inici + aligera + pagoscaja - pedidosguardadosefectivo

[Forms]![Cierre por Turno 2]![montoperiodo] = pedidosguardados
[Forms]![Cierre por Turno 2]![vposcalculado] = pedidosguardadospos
[Forms]![Cierre por Turno 2]![vtrasnfercalculado] = pedidosguardadostrasnfer
[Forms]![Cierre por Turno 2]![vpedidosyacalculado] = pedidosguardadospedidosya
[Forms]![Cierre por Turno 2]![vefeccalculado] = pedidosguardadosefectivo


' antigua impresion de reportes'''

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE CierreDiario SET CierreDiario.MFCor = " & gcordoba & ", " & _
             "CierreDiario.MFDol = " & gdola & ", " & _
             "CierreDiario.Faltante = " & gsobrafalta & ", " & _
             "CierreDiario.TotalPedidosYa = " & gventaspya & ", " & _
             "CierreDiario.TotalTransferencia = " & gventastrans & ", " & _
             "CierreDiario.TotalPOS = " & gventaspos & " " & _
             "WHERE CierreDiario.CodigoCierre = " & dcierre & ""
DoCmd.SetWarnings True

Call SyncCierreDiario30Dias

DoCmd.OpenReport "Contador Caja 2", acViewNormal
Me.Requery


DoCmd.OpenForm "Cierre por Turno 3"
[Forms]![Cierre por Turno 3]![ccodigocierre] = dcierre + 1
[Forms]![Cierre por Turno 3]![ccodigocajero] = dopera
[Forms]![Cierre por Turno 3]![cfecha] = dfecha
[Forms]![Cierre por Turno 3]![cinicio] = dinicio
[Forms]![Cierre por Turno 3]![cfinal] = dfinal

[Forms]![Cierre por Turno 3]![posguardado] = gventaspos
[Forms]![Cierre por Turno 3]![transguardado] = gventastrans
[Forms]![Cierre por Turno 3]![pyaguardado] = gventaspya

[Forms]![Cierre por Turno 3]![totalcordobas] = gcordoba
[Forms]![Cierre por Turno 3]![totaldolares] = gdola

Call Forms("Cierre por Turno 3").actualizariconoscierre

DoCmd.Close acForm, "Contador Caja 2"
End Sub




Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.productoscomprassistema.height = 280 * CantidadProductosCompradosCajaDia(Date)
Me.InsideHeight = Me.InsideHeight + 280 * CantidadProductosCompradosCajaDia(Date)
End Sub

Private Sub c0_5_Exit(Cancel As Integer)
If IsNull(Me.[c0.5]) Then
    Me.[c0.5] = 0
End If
End Sub

Private Sub c1_Exit(Cancel As Integer)
If IsNull(Me.c1) Then
    Me.c1 = 0
End If
End Sub

Private Sub c5_Exit(Cancel As Integer)
If IsNull(Me.c5) Then
    Me.c5 = 0
End If
End Sub

Private Sub c10_Exit(Cancel As Integer)
If IsNull(Me.c10) Then
    Me.c10 = 0
End If
End Sub

Private Sub c20_Exit(Cancel As Integer)
If IsNull(Me.c20) Then
    Me.c20 = 0
End If
End Sub

Private Sub c50_Exit(Cancel As Integer)
If IsNull(Me.c50) Then
    Me.c50 = 0
End If
End Sub

Private Sub c100_Exit(Cancel As Integer)
If IsNull(Me.c100) Then
    Me.c100 = 0
End If
End Sub

Private Sub c200_Exit(Cancel As Integer)
If IsNull(Me.c200) Then
    Me.c200 = 0
End If
End Sub

Private Sub c500_Exit(Cancel As Integer)
If IsNull(Me.c500) Then
    Me.c500 = 0
End If
End Sub

Private Sub d1_Exit(Cancel As Integer)
If IsNull(Me.d1) Then
    Me.d1 = 0
End If
End Sub

Private Sub d5_Exit(Cancel As Integer)
If IsNull(Me.d5) Then
    Me.d5 = 0
End If
End Sub

Private Sub d10_Exit(Cancel As Integer)
If IsNull(Me.d10) Then
    Me.d10 = 0
End If
End Sub

Private Sub d20_Exit(Cancel As Integer)
If IsNull(Me.d20) Then
    Me.d20 = 0
End If
End Sub

Private Sub d50_Exit(Cancel As Integer)
If IsNull(Me.d50) Then
    Me.d50 = 0
End If
End Sub

Private Sub d100_Exit(Cancel As Integer)
If IsNull(Me.d100) Then
    Me.d100 = 0
End If
End Sub

Private Sub tventaspedidosya_Exit(Cancel As Integer)

If IsNull(Me.tventaspedidosya) Then
    Me.tventaspedidosya = 0
End If


End Sub

Private Sub tventasposs_Exit(Cancel As Integer)

If IsNull(Me.tventasposs) Then
    Me.tventasposs = 0
End If


End Sub

Private Sub tventastransferencia_Exit(Cancel As Integer)

If IsNull(Me.tventastransferencia) Then
    Me.tventastransferencia = 0
End If


End Sub
