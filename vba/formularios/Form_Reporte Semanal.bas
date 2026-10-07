' ==========================================================
' Modulo  : Form_Reporte Semanal
' Tipo    : 100  |  Lineas: 225
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database

Public Sub Comando175_Click()
If codigoLocal() = 0 Then
    Exit Sub
End If
    
'Me.acierre.Height = Me.alturadepositos * 280
Me.ventasdiarias.height = 7 * 280
Me.asemana.height = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.asemana & " - 1.Reporte Semanal.pdf"
DoCmd.OutputTo acOutputForm, "Reporte Semanal", acFormatPDF, archivo, False
End Sub

Public Sub Comando56_Click()

Me.titulo.Caption = "REPORTE SEMANA " & Me.asemana & " - " & UCase(ciudadsistema())

'========================Ventas Reales===================================
'Ventas semanal, dia por dia
Dim rst As DAO.Recordset
Dim miSQL As String
Dim tempfecha As Date
Dim tempvt, tempcierre, tempcajainicial, tempcomprascaja, tempventaspos, tempfaltantes As Double

miSQL = "SELECT numerosemana([FechaSistema]![Dates]) AS semana, FechaSistema.Dates" & _
" FROM FechaSistema" & _
" WHERE (((numerosemana([FechaSistema]![Dates])) = " & Me.asemana & "))" & _
" ORDER BY FechaSistema.Dates"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Me.ventasdiarias = ""
Me.listacierre = ""
Me.listacajainicial = ""
Me.listacomprascaja = ""
Me.listaventaspos = ""
Me.listafaltantes = ""

Me.listaobservaciones = ""
Me.historial = ""

Me.VT = 0 'VentaRealSemanal(Me.asemana)
Me.cierre = 0 'SumaCierresSemanaInterno(Me.asemana)
Me.cajaini = 0 'SumaCajaInicialSemanaInterno(Me.asemana)
Me.comprascaja = 0 'SalidasDeCajaSemanaInterno(Me.asemana)
Me.ventaspos = 0 'FRVentaTotalSemanaLocal(Me.asemana, 1)
Me.faltantes = 0

Me.periodo.Caption = rst("Dates") & " - "
Me.rango = "Ventas " & Me.asemana - 7 & "->" & Me.asemana - 1

For I = 1 To 7 ' 7 dias de la semana
    tempfecha = rst("Dates")
    Me.periodo.Caption = Me.periodo.Caption & IIf(I = 7, tempfecha, "")
    
    tempvt = AcumuladoDia(tempfecha)
    Me.VT = Me.VT + tempvt
    Me.ventasdiarias = Me.ventasdiarias & I & ":  " & Format(tempvt, "#,#0.00#") & Chr(13) & Chr(10)
    
    tempcierre = MontoCierre(tempfecha, "T")
    Me.cierre = Me.cierre + tempcierre
    Me.listacierre = Me.listacierre & Format(tempcierre, "#,#0.00#") & Chr(13) & Chr(10)
    
    tempcajainicial = cajainicial(tempfecha)
    Me.cajaini = Me.cajaini + tempcajainicial
    Me.listacajainicial = Me.listacajainicial & Format(tempcajainicial, "#,#0.00#") & Chr(13) & Chr(10)
    
    tempcomprascaja = pagospordia(tempfecha)
    Me.comprascaja = Me.comprascaja + tempcomprascaja
    Me.listacomprascaja = Me.listacomprascaja & Format(tempcomprascaja, "#,#0.00#") & Chr(13) & Chr(10)
    
    tempventaspos = montoposdia(tempfecha)
    Me.ventaspos = Me.ventaspos + tempventaspos
    Me.listaventaspos = Me.listaventaspos & Format(tempventaspos, "#,#0.00#") & Chr(13) & Chr(10)
    
    tempfaltantes = (tempcajainicial + tempvt - tempventaspos) - (tempcierre + tempcomprascaja)
    Me.faltantes = Me.faltantes + tempfaltantes
    Me.listafaltantes = Me.listafaltantes & Format(tempfaltantes, "#,#0.00#") & Chr(13) & Chr(10)
    
    Me.listaobservaciones = Me.listaobservaciones & Left(DLookup("[Eventos]", "EstadoInicial", "[Fecha] = #" & tempfecha & "#"), 20) & Chr(13) & Chr(10)
    Me.historial = Me.historial & Format(VentaRealSemanal(Me.asemana - 8 + I), "#,#0.0#") & IIf(I = 7, "", "  ->  ")
    
    rst.MoveNext
Next I
rst.Close

Me.ventasreales = Me.cierre - Me.cajaini + Me.comprascaja + Me.ventaspos

'====================Costo de Ventas Reales==================
Me.consumoteorico = FRCostoConsumoTeoricoSemana(Me.asemana)
Me.inventarioinicial = FRCostoIFCotizacion(Me.asemana - 1) + FRCostoIFIngrediente(Me.asemana - 1)
Me.inventariofinal = FRCostoIFCotizacion(Me.asemana) + FRCostoIFIngrediente(Me.asemana)
Me.Mermas = FRCostoMermasCotizacionSemana(Me.asemana) + FRCostoMermasIngredienteSemana(Me.asemana)
Me.ingresos = FRCostoIngresosSemana(Me.asemana)
Me.consumoreal = Me.inventarioinicial + Me.ingresos - Me.inventariofinal - Me.Mermas

'====================Ventas por Grupo==========================================
Me.Requery
Me.Filter = "[Cantidad] <> 0"
Me.FilterOn = True
Me.Requery

'==========================Desempeño==========================================
Me.TT = TotalTardanzas(Me.asemana)
Me.CO = ProductosPromocionSemana(6, Me.asemana)
Me.TRI = ProductosPromocionSemana(67, Me.asemana)
Me.canjeadoxpuntos = ProductosPromocionSemana(22, Me.asemana)
Me.TV = MembresiaVentaSinPromocion(Me.asemana)
Me.promovolantes = ProductosPromocionSemana(64, Me.asemana)

'==========================Perdidas==========================================
Me.DE = ProductosPromocionSemana(43, Me.asemana) + ProductosPromocionSemana(39, Me.asemana)
Me.MDV = CosteoPedidosDevueltos(Me.asemana)
Me.MR = ValorizacionMermasSemanaTotal(Me.asemana)
Me.MV = FRVariacionTotal(Me.asemana)
Me.PI = PagoInfluencers(Me.asemana)
Me.DPR = VentaTeoricoSemanal(Me.asemana) - Me.VT

'==========================Depositos==========================================
'Depositos semana vs cierre
'Dim rst2 As DAO.Recordset
'Dim miSQL2 As String
'Dim cantidaddepositos As Integer
'Dim acumuladocierre As Double
'Dim bcierre As Double
'Dim acumuladodeposito As Double
'Dim bdepositocor, bdepositodol As Double
'Dim fechaanterior As Date
'miSQL2 = "SELECT numerosemana([Depositos]![Fecha]) AS semana, Depositos.Fecha, Depositos.Observacion, FechaDepositoAnterior([Depositos]![Fecha]) AS Desde, [Depositos]![Fecha]-1 AS Hasta, Depositos.Monto, Depositos.Denominacion" & _
'" FROM Depositos" & _
'" WHERE (((numerosemana([Depositos]![Fecha]))=" & Me.asemana & "))" & _
'" ORDER BY Depositos.Fecha DESC"
'Set rst2 = CurrentDb.OpenRecordset(miSQL2, dbOpenDynaset)
'rst2.MoveLast
'cantidaddepositos = rst2.RecordCount
'rst2.MoveFirst
'Me.alturadepositos = cantidaddepositos
'Me.adesde = ""
'Me.ahasta = ""
'Me.acierre = ""
'Me.adepositocor = ""
'Me.adepositodol = ""
'Me.fdeposito = ""
'Me.Observaciones = ""
'acumuladocierre = 0
'acumuladodeposito = 0
'
'For I = 1 To cantidaddepositos
'
'    Me.adesde = Me.adesde & rst2("Desde") & Chr(13) & Chr(10)
'    Me.ahasta = Me.ahasta & rst2("Hasta") & Chr(13) & Chr(10)
'
'    bcierre = AcumuladoCierresDiarios(rst2("Desde"), rst2("Hasta"))
'
'    If I = 1 Then
'       acumuladocierre = acumuladocierre + bcierre
'       Me.acierre = Me.acierre & Format(bcierre, "#,#0.00#") & Chr(13) & Chr(10)
'    Else
'       If Not rst2("Fecha") = fechaanterior Then
'          acumuladocierre = acumuladocierre + bcierre
'          Me.acierre = Me.acierre & Format(bcierre, "#,#0.00#") & Chr(13) & Chr(10)
'       Else
'          Me.acierre = Me.acierre & Chr(13) & Chr(10)
'       End If
'    End If
'
'
'    bdepositocor = IIf(rst2("Denominacion") = "Cordobas", rst2("Monto"), 0)
'    bdepositodol = IIf(rst2("Denominacion") = "Dolares", rst2("Monto"), 0)
'    acumuladodeposito = acumuladodeposito + bdepositocor + bdepositodol * tipocambio(rst2("Fecha"))
'    Me.adepositocor = Me.adepositocor & Format(bdepositocor, "#,#0.00#") & Chr(13) & Chr(10)
'    Me.adepositodol = Me.adepositodol & Format(bdepositodol, "#,#0.00#") & Chr(13) & Chr(10)
'
'    Me.fdeposito = Me.fdeposito & rst2("Fecha") & Chr(13) & Chr(10)
'    Me.Observaciones = Me.Observaciones & " " & rst2("Observacion") & Chr(13) & Chr(10)
'    fechaanterior = rst2("Fecha")
'    rst2.MoveNext
'Next I
'Me.totalcierre = acumuladocierre
'Me.totaldeposito = acumuladodeposito
'
'If acumuladocierre - acumuladodeposito > 0 Then
'    Me.titulofaltasobra = "FALTA"
'    Me.faltasobra = Format(acumuladocierre - acumuladodeposito, "#,#0.0#")
'Else
'    Me.titulofaltasobra = "SOBRA"
'    Me.faltasobra = Format(acumuladodeposito - acumuladocierre, "#,#0.0#")
'End If

'rst2.Close

'==========================Costos==========================================

Me.variacionventas = IIf(Me.VT > Me.ventasreales, "(-) " & Format(Me.VT - Me.ventasreales, "#,#0.00#"), "(+) " & Format(Me.ventasreales - Me.VT, "#,#0.00#"))
Me.variacioncostoventas = IIf(Me.consumoreal > Me.consumoteorico, "(-) " & Format(Me.consumoreal - Me.consumoteorico, "#,#0.00#"), "(+) " & Format(Me.consumoteorico - Me.consumoreal, "#,#0.00#"))
Me.margenbruto = Me.ventasreales - Me.consumoreal
Me.margenbrutoventas = Me.margenbruto / Me.ventasreales
Me.ventasposventas = Me.ventaspos / Me.ventasreales
Me.promoventassinpromo = (FRVentaTotalSemanaLocal(Me.asemana, 0) - ProductosPromocionSemana(5, Me.asemana) - ProductosPromocionSemana(43, Me.asemana) - ProductosPromocionSemana(39, Me.asemana) - ProductosPromocionSemana(6, Me.asemana)) / FRVentaTotalSemanaLocal(Me.asemana, 0)
End Sub


Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
