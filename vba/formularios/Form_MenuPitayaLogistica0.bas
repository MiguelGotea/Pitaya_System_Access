' ==========================================================
' Modulo  : Form_MenuPitayaLogistica0
' Tipo    : 100  |  Lineas: 403
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:16
' ==========================================================

Option Compare Database


Private Function generarlistasecos(fechi As Date) As Long

If Me.listasecos = 0 Or IsNull(Me.listasecos) Then
    generarlistasecos = crearpreingreso(fechi, Me.relocal)
    Me.listasecos = generarlistasecos
Else
    generarlistasecos = Me.listasecos
End If
End Function
Private Function generarlistacongelados(fechi As Date) As Long
If Me.listacongelados = 0 Or IsNull(Me.listacongelados) Then
    generarlistacongelados = crearpreingreso(fechi, Me.relocal)
    Me.listacongelados = generarlistacongelados
Else
    generarlistacongelados = Me.listacongelados
End If
End Function

Private Function generarlistamostrador(fechi As Date) As Long
If Me.listamostrador = 0 Or IsNull(Me.listamostrador) Then
    generarlistamostrador = crearpreingreso(fechi, Me.relocal)
    Me.listamostrador = generarlistamostrador
Else
    generarlistamostrador = Me.listamostrador
End If
End Function



Private Sub Comando1217_Click()
DoCmd.OpenForm "Introduccion"
DoCmd.Close acForm, "MenuPitayaLogistica0"
End Sub


Private Sub Comando1425_Click()
DoCmd.OpenForm "CompraProcesamientoSemana"
End Sub

Private Sub Comando1429_Click()
DoCmd.OpenForm "Calculo CU Cotizacion Semana NP"
End Sub

Private Sub Comando1430_Click()
DoCmd.OpenForm "Calculo Conversion Cotizacion Semana"
End Sub

Private Sub Comando1431_Click()
DoCmd.OpenForm "Calculo CU Ingrediente Semana"
End Sub

Private Sub Comando1436_Click()
DoCmd.OpenForm "Calculo CU Cotizacion Semana P"
End Sub

Private Sub Comando1438_Click()
DoCmd.OpenForm "Calculo Pareto Semana"
End Sub

Private Sub Comando1474_Click()
DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").modocentral
End Sub

Private Sub Comando1490_Click()
DoCmd.OpenForm "SeguimientoPlanPorciones"
End Sub

Private Sub Comando1505_Click()
DoCmd.OpenForm "CompraxProovedor"
End Sub



Private Sub Comando1510_Click()
Call actualizartablasmixedglobal

End Sub







Private Sub Comando1557_Click()
DoCmd.OpenForm "ConsumoNoPorciones"
End Sub


Private Sub Comando1559_Click()
DoCmd.OpenForm "PlanProduccionMarcaPitaya"
End Sub

Private Sub Comando1565_Click()

'Porciones
DoCmd.OpenForm "SeguimientoPlanPorciones"
[Forms]![SeguimientoPlanPorciones].asemana = Me.resemana
[Forms]![SeguimientoPlanPorciones].rango = 4
[Forms]![SeguimientoPlanPorciones].asucursal = Me.relocal
[Forms]![SeguimientoPlanPorciones].incre = Me.reincre
Call Forms("[SeguimientoPlanPorciones]").asucursal_Exit(0)
Call Forms("[SeguimientoPlanPorciones]").Comando1040_Click

' Productos NO Porciones No Perecibles
DoCmd.OpenForm "ConsumoNoPorciones"
[Forms]![ConsumoNoPorciones].asemana = Me.resemana
[Forms]![ConsumoNoPorciones].crango = 4
[Forms]![ConsumoNoPorciones].alocali = Me.relocal
[Forms]![ConsumoNoPorciones].incre = Me.reincre
Call Forms("[ConsumoNoPorciones]").alocali_Exit(0)
Call Forms("[ConsumoNoPorciones]").Comando170_Click

'Productos Mostrador
DoCmd.OpenForm "ConsumoPitayaStore"
[Forms]![ConsumoPitayaStore].asemana = Me.resemana
[Forms]![ConsumoPitayaStore].arango = 4
[Forms]![ConsumoPitayaStore].alocali = Me.relocal
[Forms]![ConsumoPitayaStore].incre = Me.reincre
Call Forms("[ConsumoPitayaStore]").alocali_Exit(0)
Call Forms("[ConsumoPitayaStore]").Comando170_Click



End Sub

Private Sub Comando1570_Click()
DoCmd.OpenForm "PlanPorcionesSemanaTotal"
End Sub

Private Sub Comando1917_Click()
DoCmd.OpenForm "PlanDespachoInsumosFijosConsumibles"
End Sub

Private Sub Comando1977_Click()
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.Nombre, DatosSistema.IngresoInsumos, DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0) And ((DatosSistema.CodSistema)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    Me.relocal = rst("CodSistema")
    Call Comando1565_Click
    rst.MoveNext
Loop
rst.Close
Exit Sub

Nulo:
MsgBox "Error en la generaciond e listas"

End Sub

Private Sub Comando1630_Click()
DoCmd.OpenForm "ConsumoPitayaStore"
End Sub

Private Sub Comando1674_Click()
Call eliminartablasmain
Call importartablasmain

Call eliminartablascentral(codigoLocal())
Call importartablascentral(codigoLocal())

Call importartablasweb
MsgBox "Tablas Main Actualizadas"
End Sub

Private Sub Comando1679_Click()
DoCmd.OpenForm "ResumenPreIngresosPitaya"
[Forms]![ResumenPreIngresosPitaya]![aencabezado].Caption = "RESUMEN DE INGRESOS DE LA SEMANA"
End Sub

Private Sub Comando1702_Click()
Call descargardatoscentralcopia("todos")
MsgBox "Datos de la central completos"
End Sub


Private Sub Comando1715_Click()
DoCmd.OpenForm "HistorialResumenPagos"
End Sub

Private Sub Comando1729_Click()
DoCmd.OpenForm "HistorialOrdenDeCompra"
End Sub

Private Sub Comando1738_Click()
DoCmd.OpenForm "GenerarIngresoDeCompras"
End Sub

Private Sub Comando1750_Click()
DoCmd.OpenForm "ResumenRegistroPreIngresosPitaya"
End Sub


Private Sub Comando1762_Click()
DoCmd.OpenForm "ComprasSemanalesRefrigerados"
End Sub

Private Sub Comando1904_Click()
DoCmd.OpenForm "HistorialDespachoInsumosFIjos"
End Sub

Private Sub Comando1915_Click()
Dim codpre As Long
codpre = crearpreingreso(Me.Fecha2, Me.relocal)
Call AutoIngresoDatosInsumosFijos(Me.resemana, Me.relocal, codpre)

End Sub





Private Sub Comando1979_Click()
DoCmd.OpenForm "ComprasSemanalesSecos"
End Sub

Private Sub Comando2003_Click()
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.IngresoInsumos, DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0) And ((DatosSistema.CodSistema)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    Me.relocal = rst("CodSistema")
    Me.Fecha1 = FechaDeNumeroDiaSemana(Me.resemana, rst("IngresoInsumos"))
    Me.Fecha2 = Me.Fecha1 + 3
    Call Comando1898_Click
    rst.MoveNext
Loop
rst.Close
Exit Sub

Nulo:
MsgBox "Error al generar preingresos"
End Sub

Private Sub Comando2006_Click()
Dim semanab As Integer
semanab = Me.resemana
DoCmd.OpenForm "ResumenRegistroPreIngresosPitaya"

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.Nombre, DatosSistema.IngresoInsumos, DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0) And ((DatosSistema.CodSistema)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    [Forms]![ResumenRegistroPreIngresosPitaya]![sucursalac] = rst("Nombre")
    
    'Primer Despacho
    [Forms]![ResumenRegistroPreIngresosPitaya]![fechaac] = FechaDeNumeroDiaSemana(semanab, rst("IngresoInsumos"))
    Call Forms("[ResumenRegistroPreIngresosPitaya]").Comando214_Click
    
    'Degundo Despacho
    [Forms]![ResumenRegistroPreIngresosPitaya]![fechaac] = [Forms]![ResumenRegistroPreIngresosPitaya]![fechaac] + 3
    Call Forms("[ResumenRegistroPreIngresosPitaya]").Comando214_Click
    
    rst.MoveNext
Loop
rst.Close
Exit Sub

Nulo:
MsgBox "PDF Generados"
End Sub

Private Sub Comando2047_Click()
DoCmd.OpenForm "PlanProduccionMarcaPitaya"
[Forms]![PlanProduccionMarcaPitaya].asemana = [Forms]![MenuPitayaProduccion0].semanaplanproduccion
Call Forms("[PlanProduccionMarcaPitaya]").asemana_Exit(0)
Call Forms("[PlanProduccionMarcaPitaya]").Comando170_Click

DoCmd.OpenForm "PlanPorcionesSemanaTotal"
[Forms]![PlanPorcionesSemanaTotal].asemana = [Forms]![MenuPitayaProduccion0].semanaplanproduccion
Call Forms("[PlanPorcionesSemanaTotal]").asemana_Exit(0)
Call Forms("[PlanPorcionesSemanaTotal]").Comando1040_Click

End Sub

Private Sub Comando2058_Click()
DoCmd.OpenForm "CompraInsumosFijosConsumiblesGlobal"
End Sub

Private Sub Comando2075_Click()
DoCmd.OpenForm "RevisionInventariosCargadosSucursalesVariables"
End Sub

Private Sub Comando2085_Click()
DoCmd.OpenForm "RevisionInventariosCargadosSucursalesFijos"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "BATIDOS PITAYA 0"

End Sub

Private Sub Form_Timer()
On Error Resume Next
Dim Horal As Date
Horal = Time()

'Call ComunicacionTelegram

'================================================================================================
'Fotos hora programada
If Minute(Horal) = 0 And Second(Horal) < 30 Then 'Foto abierto local
    Call FotoTemporalTelegram
End If

End Sub

Private Sub Comando1829_Click()
Dim codpres As Long
Dim codprec As Long
codpres = generarlistasecos(Me.Fecha1)
codprec = generarlistacongelados(Me.Fecha1)

Call AutoIngresoDatosPorciones(Me.resemana, Me.rerango, Me.reincre, Me.relocal, 1, codpres, codprec, 1)
End Sub

Private Sub Comando1839_Click()
Dim codpres As Long
Dim codprec As Long
codpres = generarlistasecos(Me.Fecha2)
codprec = generarlistacongelados(Me.Fecha2)
Call AutoIngresoDatosPorciones(Me.resemana, Me.rerango, Me.reincre, Me.relocal, 2, codpres, codprec, 1)
End Sub

Private Sub Comando1841_Click()
Dim codpres As Long
Dim codprec As Long
codpres = generarlistasecos(Me.Fecha1)
codprec = generarlistacongelados(Me.Fecha1)
Call AutoIngresoDatosNoPerecibles(Me.resemana, Me.rerango, Me.reincre, Me.relocal, 1, codpres, codprec, 1)
End Sub

Private Sub Comando1843_Click()
Dim codpres As Long
Dim codprec As Long
codpres = generarlistasecos(Me.Fecha2)
codprec = generarlistacongelados(Me.Fecha2)
Call AutoIngresoDatosNoPerecibles(Me.resemana, Me.rerango, Me.reincre, Me.relocal, 2, codpres, codprec, 1)
End Sub
Private Sub Comando1845_Click()
Dim codpres As Long
Dim codprec As Long
codpres = generarlistamostrador(Me.Fecha1)
codprec = codpres
Call AutoIngresoDatosMostrador(Me.resemana, Me.rerango, Me.reincre, Me.relocal, 1, codpres, codprec, 1)
End Sub

Private Sub Comando1847_Click()
Dim codpres As Long
Dim codprec As Long
codpres = generarlistamostrador(Me.Fecha2)
codprec = codpres
Call AutoIngresoDatosMostrador(Me.resemana, Me.rerango, Me.reincre, Me.relocal, 2, codpres, codprec, 1)
End Sub

Private Sub Comando1849_Click()
Call Comando1829_Click
Call Comando1841_Click
Call Comando1845_Click

Me.listacongelados = 0
Me.listasecos = 0
Me.listamostrador = 0
End Sub

Private Sub Comando1878_Click()
Call Comando1839_Click
Call Comando1843_Click
Call Comando1847_Click

Me.listacongelados = 0
Me.listasecos = 0
Me.listamostrador = 0
End Sub

Private Sub Comando1898_Click()
Call Comando1849_Click
Call Comando1878_Click
'MsgBox "Listas del Despacho Generadas"
End Sub
