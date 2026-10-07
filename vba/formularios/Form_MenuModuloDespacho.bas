' ==========================================================
' Modulo  : Form_MenuModuloDespacho
' Tipo    : 100  |  Lineas: 540
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database
Private Sub cargarfechasdespacho()

If IsNull(Me.relocal) Then
    Exit Sub
End If

If IsNull(Me.resemana) Then
    Exit Sub
End If

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.IngresoInsumos, DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0) And ((DatosSistema.CodSistema)=" & Me.relocal & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Me.Fecha1 = FechaDeNumeroDiaSemana(Me.resemana, rst("IngresoInsumos"))
Me.Fecha2 = Me.Fecha1 + 3

rst.Close
Exit Sub

Nulo:
MsgBox "Error al generar preingresos"
End Sub

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
DoCmd.Quit
End Sub



Private Sub Comando1490_Click()
DoCmd.OpenForm "SeguimientoPlanPorciones"
End Sub




Private Sub Comando1557_Click()
DoCmd.OpenForm "ConsumoNoPorciones"
End Sub

Private Sub Comando1565_Click()

'Porciones
DoCmd.OpenForm "SeguimientoPlanPorciones"
[Forms]![SeguimientoPlanPorciones].asemana = Me.resemanap
[Forms]![SeguimientoPlanPorciones].rango = Me.rerangop
[Forms]![SeguimientoPlanPorciones].asucursal = Me.relocalp
[Forms]![SeguimientoPlanPorciones].incre = Me.reincrep
Call Forms("[SeguimientoPlanPorciones]").asucursal_Exit(0)
Call Forms("[SeguimientoPlanPorciones]").Comando1040_Click

' Productos NO Porciones No Perecibles
DoCmd.OpenForm "ConsumoNoPorciones"
[Forms]![ConsumoNoPorciones].asemana = Me.resemanap
[Forms]![ConsumoNoPorciones].crango = Me.rerangop
[Forms]![ConsumoNoPorciones].alocali = Me.relocalp
[Forms]![ConsumoNoPorciones].incre = Me.reincrep
Call Forms("[ConsumoNoPorciones]").alocali_Exit(0)
Call Forms("[ConsumoNoPorciones]").Comando170_Click

'Productos Mostrador
DoCmd.OpenForm "ConsumoPitayaStore"
[Forms]![ConsumoPitayaStore].asemana = Me.resemanap
[Forms]![ConsumoPitayaStore].arango = Me.rerangop
[Forms]![ConsumoPitayaStore].alocali = Me.relocalp
[Forms]![ConsumoPitayaStore].incre = Me.reincrep
Call Forms("[ConsumoPitayaStore]").alocali_Exit(0)
Call Forms("[ConsumoPitayaStore]").Comando170_Click

End Sub



Private Sub Comando1917_Click()
DoCmd.OpenForm "PlanDespachoInsumosFijosConsumibles"
End Sub

Private Sub Comando1977_Click()
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo" & _
" FROM StatusSucursales WHERE ((Not (StatusSucursales.CodLocal)=0) AND ((StatusSucursales.Activo)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    Me.relocalp = rst("CodLocal")
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

MsgBox "Tablas Main Actualizadas"
End Sub



Private Sub Comando1702_Click()
Call descargardatoscentralcopia("despacho")
MsgBox "Datos de la central completos"
End Sub


Private Sub Comando1915_Click()
Dim codpre As Long
codpre = crearpreingreso(Me.Fecha2, Me.relocal)
Me.listafijos = codpre
Call AutoIngresoDatosInsumosFijos(Me.resemana, Me.relocal, codpre)
Me.listafijos = 0

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
    Call Comando1849_Click
    Call Comando1878_Click
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

Private Sub Comando2090_Click()
DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").mododespacho
End Sub

Private Sub Comando2106_Click()
DoCmd.OpenForm "HistorialDespachoInsumosFIjos"
End Sub

Private Sub Comando2115_Click()
DoCmd.OpenForm "ResumenPreIngresosPitaya"
[Forms]![ResumenPreIngresosPitaya]![aencabezado].Caption = "RESUMEN DE INGRESOS DE LA SEMANA"
End Sub

Private Sub Comando2116_Click()
DoCmd.OpenForm "ResumenRegistroPreIngresosPitaya"
End Sub


Private Sub Comando2620_Click()
DoCmd.OpenForm "ControlHorasArea"
End Sub

Private Sub Comando2625_Click()
DoCmd.OpenForm "HistorialOrdenDeCompra"
End Sub

Private Sub Comando2628_Click()
DoCmd.OpenForm "ResumenSemanalProduccion"
End Sub

Private Sub Comando2631_Click()
Dim semb As Integer
semb = InputBox("Ingrese el numero de semana a buscar", "Numero de Semana")
DoCmd.OpenForm "Busqueda Compras", , , "[DBIngredientes].[Tipo]='Transporte' AND [semana]=" & semb
End Sub

Private Sub Comando2634_Click()
DoCmd.OpenForm "HistorialCostoUnitarioIngredientes"
End Sub





Private Sub Comando2675_Click()
Call Comando1849_Click
Call Comando1878_Click

'Pasar datos a otra epstana
Me.resemanap = Me.resemana
Me.rerangop = Me.rerango
Me.relocalp = Me.relocal
Me.reincrep = Me.reincre

Call Comando1565_Click

MsgBox "Preingresos y Listas del Despacho Generadas"
End Sub

Private Sub Comando2677_Click()
DoCmd.OpenForm "SolicitudDespachoSucursales"
End Sub

Private Sub Comando2685_Click()
DoCmd.SetWarnings False
DoCmd.CopyObject , "CambiosPreingresoSucursal", acTable, "CambiosPreIngresosPitaya"
DoCmd.RunSQL "ALTER TABLE CambiosPreingresoSucursal ADD COLUMN Sucursal Integer"
DoCmd.RunSQL "DELETE * FROM CambiosPreingresoSucursal"

DoCmd.CopyObject , "CambiosPreingresoSucursalMixed", acTable, "CambiosPreingresoSucursal"
DoCmd.RunSQL "ALTER TABLE CambiosPreingresoSucursalMixed DROP CONSTRAINT PrimaryKey"
DoCmd.SetWarnings True



On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim sucu As Integer

miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo" & _
" FROM StatusSucursales WHERE ((Not (StatusSucursales.CodLocal)=0) AND ((StatusSucursales.Activo)<>0))" & _
" ORDER BY StatusSucursales.CodLocal"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    sucu = rst("CodLocal")
    Call importartablaespecifica("Pitaya" & sucu & "_DB", "CambiosPreIngresosPitaya", "CambiosPreingresoSucursal", 1)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE CambiosPreingresoSucursal SET CambiosPreingresoSucursal.Sucursal = " & sucu


    DoCmd.RunSQL "INSERT INTO CambiosPreingresoSucursalMixed SELECT * FROM CambiosPreingresoSucursal"
    DoCmd.RunSQL "DELETE * FROM CambiosPreingresoSucursal"
    DoCmd.SetWarnings True
    
    rst.MoveNext
Loop
rst.Close
DoCmd.OpenForm "CambiosGlobalRegistrados"

Exit Sub

Nulo:
MsgBox "Error al cargar datos"




End Sub

Private Sub Comando2689_Click()
DoCmd.OpenForm "HistoricoConsumoPorcionesDescargado"
End Sub

Private Sub Comando2693_Click()
DoCmd.OpenForm "HistoricoConsumoNoPorcionesDescargado"
End Sub

Private Sub Comando2716_Click()
DoCmd.OpenForm "HistoricoConsumoMostradorDescargado"
End Sub

Private Sub Comando2723_Click()

DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "Despacho"

End Sub

Private Sub Comando2727_Click()
DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").mododespacho
End Sub

Private Sub Comando2733_Click()
DoCmd.OpenForm "SolicitudDespachoSucursales"
End Sub

Private Sub Comando2739_Click()
DoCmd.SetWarnings False
DoCmd.CopyObject , "CambiosPreingresoSucursal", acTable, "CambiosPreIngresosPitaya"
DoCmd.RunSQL "ALTER TABLE CambiosPreingresoSucursal ADD COLUMN Sucursal Integer"
DoCmd.RunSQL "DELETE * FROM CambiosPreingresoSucursal"

DoCmd.CopyObject , "CambiosPreingresoSucursalMixed", acTable, "CambiosPreingresoSucursal"
DoCmd.RunSQL "ALTER TABLE CambiosPreingresoSucursalMixed DROP CONSTRAINT PrimaryKey"
DoCmd.SetWarnings True



'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim sucu As Integer

miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo" & _
" FROM StatusSucursales WHERE ((Not (StatusSucursales.CodLocal)=0) AND ((StatusSucursales.Activo)<>0))" & _
" ORDER BY StatusSucursales.CodLocal"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    sucu = rst("CodLocal")
    Call importartablaespecifica("Pitaya" & sucu & "_DB", "CambiosPreIngresosPitaya", "CambiosPreingresoSucursal", 1)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE CambiosPreingresoSucursal SET CambiosPreingresoSucursal.Sucursal = " & sucu


    DoCmd.RunSQL "INSERT INTO CambiosPreingresoSucursalMixed SELECT * FROM CambiosPreingresoSucursal"
    DoCmd.RunSQL "DELETE * FROM CambiosPreingresoSucursal"
    DoCmd.SetWarnings True
    
    rst.MoveNext
Loop
rst.Close
DoCmd.OpenForm "CambiosGlobalRegistrados"

Exit Sub

Nulo:
MsgBox "Error al cargar datos"

End Sub



Private Sub Comando2868_Click()
DoCmd.OpenForm "HistoricoConsumoPorcionesDescargado"
End Sub

Private Sub Comando2869_Click()
DoCmd.OpenForm "HistoricoConsumoNoPorcionesDescargado"
End Sub

Private Sub Comando2870_Click()
DoCmd.OpenForm "HistoricoConsumoMostradorDescargado"
End Sub

Private Sub Comando2889_Click()
DoCmd.OpenForm "HistoricoConsumoIngredienteDescargado"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "BATIDOS PITAYA 0"
Me.ShortcutMenu = False

End Sub

Private Sub Form_Timer()
On Error Resume Next
Dim Horal As Date
Horal = Time()

'Call ComunicacionTelegram

'================================================================================================
'Fotos hora programada
'If Minute(Horal) = 0 And Second(Horal) < 30 Then 'Foto abierto local
'    Call FotoTemporalTelegram
'End If

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
MsgBox "Preingresos del Despacho Generadas"
End Sub


Private Sub Comando2144_Click()
Call Comando1977_Click
Call Comando2003_Click
Call Comando2006_Click

MsgBox "Plan Completo de la semana"
End Sub

Private Sub relocal_Exit(Cancel As Integer)
Call cargarfechasdespacho
End Sub

Private Sub resemana_Exit(Cancel As Integer)
Call cargarfechasdespacho

End Sub

