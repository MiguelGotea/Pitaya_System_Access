' ==========================================================
' Modulo  : Form_Menu Gestion
' Tipo    : 100
' Lineas  : 1940
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:20
' ==========================================================
Option Compare Database
Option Explicit

Sub funcionrotacionsucursalesactivas(boton As String)
'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim locax As Integer
Dim contax As Integer
Dim contloca As Integer
Dim subconta As Integer


miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE (((StatusSucursales.Activo)<>0) AND ((StatusSucursales.Sucursal)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
contax = rst.RecordCount
rst.MoveFirst

For contloca = 1 To contax

    miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
    " FROM StatusSucursales WHERE (((StatusSucursales.Activo)<>0) AND ((StatusSucursales.Sucursal)<>0))"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    
    For subconta = 1 To contax
        If subconta = contloca Then
            locax = rst("CodLocal")
        End If
        rst.MoveNext
    Next subconta
    
    rst.Close
    

    If CurrentProject.Name = "Pitaya_System.accdb" Or CurrentProject.Name = "Pitaya_System - Copy.accdb" Or CurrentProject.Name = "Pitaya_System - Copia.accdb" Then 'archivo base de sistema
            Call actualizartablas(1, locax)
    Else
        MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
        Exit Sub
    End If


    Dim localMe As Object
    Set localMe = Me
    CallByName localMe, boton, VbMethod ', arguments would go here
        
Next contloca

MsgBox "Archivos de Excel generados correctamente"
rst.Close
Exit Sub

Nulo:
MsgBox "Problemas al egnerar archivos de Excel"
End Sub
    
Private Sub Comando1063_Click()
DoCmd.OpenForm "WebBrowser"
End Sub



Private Sub Comando1088_Click()
DoCmd.OpenForm "Relacion Pedidos Dia"
[Forms]![Relacion Pedidos Dia]![fechaactual] = Date
End Sub

Private Sub Comando1427_Click()
DoCmd.OpenForm "Relacion de Productos Venta"
End Sub

Private Sub Comando1474_Click()
'Call importarpreingresossistema0
Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)

DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").modosucursaladministracion
End Sub

Private Sub Comando1547_Click()
DoCmd.OpenForm "Control Mensual Existencias Fijos"
End Sub

Private Sub Comando1217_Click()

DoCmd.OpenForm "Introduccion"
DoCmd.Close acForm, "Menu Gestion"
End Sub



Private Sub Comando1306_Click()
DoCmd.OpenForm "Calculo Consumo Porcion Semana"
End Sub

Private Sub Comando1354_Click()
DoCmd.OpenForm "Calculo CU Cotizacion Semana NP"
End Sub

Private Sub Comando1356_Click()
DoCmd.OpenForm "Calculo Conversion Cotizacion Semana"
End Sub

Private Sub Comando1373_Click()
DoCmd.OpenForm "Calculo Pareto Semana"
End Sub





Private Sub Comando1396_Click()
DoCmd.OpenForm "EleccionSucursalVigente"
[Forms]![EleccionSucursalVigente].tipocarga = 2


End Sub



Private Sub Comando1592_Click()
Call actualizartablasmixedglobal
End Sub

Private Sub Comando1412_Click()
DoCmd.OpenForm "Control Porcionamiento"

End Sub





Private Sub Comando1916_Click()
DoCmd.OpenForm "Relacion de Productos"
End Sub

Private Sub Comando2122_Click()
DoCmd.OpenForm "ResumenVentasMensual"
End Sub

Private Sub Comando2130_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anhasta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anhasta & "\" & Me.mehasta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''VENTAS MENSUAL GRUPO
archivo = Direccionmes & "\P" & codigoLocal() & " " & Me.mehasta & "_" & Me.anhasta & " - " & DLookup("NombreGrupo", "Grupos", "[CodGrupo]=" & Me.cfgrupo) & " - Ventas Mensuales Grupos (Cantidad).jpg"
DoCmd.OpenForm "Dashboard_Ventas_MesGrupos", acFormPivotChart, , "[periodo]>=" & Me.andesde * 100 + Me.medesde & " AND [periodo]<=" & Me.anhasta * 100 + Me.mehasta & " AND [CodGrupo]=" & Me.cfgrupo
[Forms]![Dashboard_Ventas_MesGrupos].Form.Requery
[Forms]![Dashboard_Ventas_MesGrupos].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_MesGrupos"
End Sub





Private Sub Comando2142_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.atic
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.atic & "\" & Me.mtic
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''ticket promedio mensual
archivo = Direccionmes & "\P" & codigoLocal() & " - Ticket Promedio Diario.jpg"
DoCmd.OpenForm "Dashboard_Ventas_TicketDiarioPromedio", acFormPivotChart, , "[ames]=" & Me.mtic & " AND [aano]=" & Me.atic
[Forms]![Dashboard_Ventas_TicketDiarioPromedio].Form.Requery
[Forms]![Dashboard_Ventas_TicketDiarioPromedio].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_TicketDiarioPromedio"
End Sub

Private Sub Comando2179_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.acuma
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.acuma & "\" & Me.acumm
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''ticket promedio mensual
archivo = Direccionmes & "\P" & codigoLocal() & " - Acumulado Meta Mensual.jpg"
DoCmd.OpenForm "Dashboard_Ventas_DiaMesAcumulado", acFormPivotChart, , "[ames]=" & Me.acumm & " AND [aano]=" & Me.acuma & " AND [Dates]<=#" & Date & "#"
[Forms]![Dashboard_Ventas_DiaMesAcumulado].Form.Requery
[Forms]![Dashboard_Ventas_DiaMesAcumulado].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_DiaMesAcumulado"
End Sub

Private Sub Comando2202_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.diaa
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.diaa & "\" & Me.diam
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''venta diaria  mensual
archivo = Direccionmes & "\P" & codigoLocal() & " - Dia a DIa Mensual.jpg"
DoCmd.OpenForm "Dashboard_Ventas_DiaMes", acFormPivotChart, , "[ames]=" & Me.diam & " AND [aano]=" & Me.diaa & " AND [Dates]<=#" & Date & "#"
[Forms]![Dashboard_Ventas_DiaMes].Form.Requery
[Forms]![Dashboard_Ventas_DiaMes].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_DiaMes"
End Sub

Private Sub Comando2238_Click()

'''''''''''''''COMPRA VS INGRESOS''''''''''''''''''''''''''''''

DoCmd.OpenForm "Corroborar Ingresos Pitaya"
[Forms]![Corroborar Ingresos Pitaya].semanaac = Me.semareporte
Call Forms("[Corroborar Ingresos Pitaya]").Comando170_Click

'''''''''''''''PORCIONAMIENTO VS INGRESOS''''''''''''''''''''''''''''''

DoCmd.OpenForm "Porcionamiento vs Ingresos"
[Forms]![Porcionamiento vs Ingresos].asemana = Me.semareporte
Call Forms("[Porcionamiento vs Ingresos]").Comando314_Click

'''''''''''''''PORCIONAMIENTO VS INGRESOS ALNACEN''''''''''''''''''''''''''''''

DoCmd.OpenForm "Porcionamiento vs Ingresos Almacen"
[Forms]![Porcionamiento vs Ingresos Almacen].asemana = Me.semareporte
Call Forms("[Porcionamiento vs Ingresos Almacen]").Comando314_Click

'''''''''''''''PROCESAMIENTOS VS INGRESOS''''''''''''''''''''''''''''''

DoCmd.OpenForm "Control Procesamiento / Ingresos"
[Forms]![Control Procesamiento / Ingresos].asemana = Me.semareporte
Call Forms("[Control Procesamiento / Ingresos]").Comando206_Click

'''''''''''''''PROCESAMIENTO VS PORCIONES''''''''''''''''''''''''''''''

DoCmd.OpenForm "ProcesamientovsPorcionamiento"
[Forms]![ProcesamientovsPorcionamiento].asemana = Me.semareporte
Call Forms("[ProcesamientovsPorcionamiento]").Comando295_Click

'''''''''''''''CONTROL PORCIONES''''''''''''''''''''''''''''''

DoCmd.OpenForm "Control Porcionamiento"
[Forms]![Control Porcionamiento].asemana = Me.semareporte
Call Forms("[Control Porcionamiento]").Comando133_Click

'''''''''''''''CONTROL MOSTRADOR''''''''''''''''''''''''''''''

DoCmd.OpenForm "Control Existencias Compra Venta"
[Forms]![Control Existencias Compra Venta].asemana = Me.semareporte
Call Forms("[Control Existencias Compra Venta]").Comando352_Click

'''''''''''''''CONTROL NO PORCIONES''''''''''''''''''''''''''''''

DoCmd.OpenForm "ControlProductosGenerales"
[Forms]![ControlProductosGenerales].asemana = Me.semareporte
Call Forms("[ControlProductosGenerales]").Comando314_Click

'''''''''''''''RESULTADO INVENTARIOS''''''''''''''''''''''''''''''
    
DoCmd.OpenForm "Control Inventarios"
[Forms]![Control Inventarios].semana1 = Me.semareporte
Call Forms("[Control Inventarios]").Comando558_Click

'''''''''''''''CONTROL INSUMOS IMPORTANTES''''''''''''''''''''''''''''''
    
DoCmd.OpenForm "Control Insumos Importantes"
[Forms]![Control Insumos Importantes].semana1 = Me.semareporte
Call Forms("[Control Insumos Importantes]").Comando562_Click

'''''''''''''''CIERRES VS FACTURAS SEMANAL''''''''''''''''''''''''''''''
    
DoCmd.OpenForm "Control Semanal Cierres y Compras"
[Forms]![Control Semanal Cierres y Compras].asemana = Me.semareporte
Call Forms("[Control Semanal Cierres y Compras]").Comando496_Click

'''''''''''''''''''''''''''FIN'''''''''''''''''''''''''

Dim beepc As Integer
beepc = 0
Do Until beepc = 10
    beepc = beepc + 1
    DoCmd.Beep
    Sleep 1000
Loop
MsgBox "Reportes guardados en PDF"

End Sub

Private Sub Comando2304_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anohora
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anohora & "\" & Me.meshora
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''horas diaria  mensual
archivo = Direccionmes & "\P" & codigoLocal() & " - Horas Trabajadas Dia a DIa Mensual.jpg"
DoCmd.OpenForm "Dashboard_Horas_DiaMes", acFormPivotChart, , "[ames]=" & Me.meshora & " AND [aano]=" & Me.anohora & " AND [Dates]<=#" & Date & "#"
[Forms]![Dashboard_Horas_DiaMes].Form.Requery
[Forms]![Dashboard_Horas_DiaMes].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Horas_DiaMes"
End Sub

Private Sub Comando2327_Click()
Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anodiahora
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anodiahora & "\" & Me.mesdiahora
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''horas diaria  mensual vs ventas
archivo = Direccionmes & "\P" & codigoLocal() & " - Horas Trabajadas vs Ventas Totales Dia a DIa Mensual.jpg"
DoCmd.OpenForm "Dashboard_Ventas_DiaMes_Horas", acFormPivotChart, , "[ames]=" & Me.mesdiahora & " AND [aano]=" & Me.anodiahora & " AND [Dates]<=#" & Date & "#"
[Forms]![Dashboard_Ventas_DiaMes_Horas].Form.Requery
[Forms]![Dashboard_Ventas_DiaMes_Horas].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_DiaMes_Horas"
End Sub

Public Sub Comando2352_Click()
'''''''''''''''''VENTAS DIARIO x GRUPO '''''''''''''''''''''
DoCmd.OpenForm "Calculo Ventas Diarias Semana"
[Forms]![Calculo Ventas Diarias Semana].asemana = Me.semguar
Call Forms("[Calculo Ventas Diarias Semana]").Comando170_Click
DoCmd.Close acForm, "Calculo Ventas Diarias Semana"

'''''''''''''''''VENTAS DIARIO x GRUP x DELI X TIPO '''''''''''''''''''''
DoCmd.OpenForm "Calculo Ventas Diarias Tipo Semana"
[Forms]![Calculo Ventas Diarias Tipo Semana].asemana = Me.semguar
Call Forms("[Calculo Ventas Diarias Tipo Semana]").Comando170_Click
DoCmd.Close acForm, "Calculo Ventas Diarias Tipo Semana"

'''''''''''''''''CONVERSION COTIZACIONES '''''''''''''''''''''
DoCmd.OpenForm "Calculo Conversion Cotizacion Semana"
[Forms]![Calculo Conversion Cotizacion Semana].asemana = Me.semguar
Call Forms("[Calculo Conversion Cotizacion Semana]").Comando170_Click
DoCmd.Close acForm, "Calculo Conversion Cotizacion Semana"

'''''''''''''''''CONSUMO INGREDIENTES '''''''''''''''''''''
DoCmd.OpenForm "Calculo Consumo ingrediente Semana"
[Forms]![Calculo Consumo ingrediente Semana].asemana = Me.semguar
Call Forms("[Calculo Consumo ingrediente Semana]").Comando170_Click
DoCmd.Close acForm, "Calculo Consumo ingrediente Semana"

'''''''''''''''''CONSUMO PORCIONES '''''''''''''''''''''
DoCmd.OpenForm "Calculo Consumo Porcion Semana"
[Forms]![Calculo Consumo Porcion Semana].asemana = Me.semguar
Call Forms("[Calculo Consumo Porcion Semana]").Comando170_Click
DoCmd.Close acForm, "Calculo Consumo Porcion Semana"

'''''''''''''''''CONSUMO NO PORCIONES '''''''''''''''''''''
DoCmd.OpenForm "Calculo Consumo NoPorcion Semana"
[Forms]![Calculo Consumo NoPorcion Semana].asemana = Me.semguar
Call Forms("[Calculo Consumo NoPorcion Semana]").Comando170_Click
DoCmd.Close acForm, "Calculo Consumo NoPorcion Semana"

'''''''''''''''''GRAFICO VENTAS SEMANAL '''''''''''''''''''''
Me.mdesde = Me.semguar - 12
Me.mhasta = Me.semguar
Call Comando274_Click

'Call ReiniciarGoogleDrive
MsgBox "Datos Guardados"
End Sub



Private Sub Comando2402_Click()
DoCmd.OpenForm "ResumenVentasMensual"
End Sub

Public Sub Comando2413_Click()
DoCmd.OpenForm "HistorialVentasFiltro", acNormal
End Sub

Private Sub Comando2497_Click()

'''''''''''''''COMPRA VS INGRESOS''''''''''''''''''''''''''''''

DoCmd.OpenForm "Corroborar Ingresos Pitaya"
[Forms]![Corroborar Ingresos Pitaya].semanaac = Me.asemaprod
Call Forms("[Corroborar Ingresos Pitaya]").Comando170_Click

'''''''''''''''PORCIONAMIENTO VS INGRESOS''''''''''''''''''''''''''''''

DoCmd.OpenForm "PorcionamientoVsIngresosGlobal"
[Forms]![PorcionamientoVsIngresosGlobal].asemana = Me.asemaprod
Call Forms("[PorcionamientoVsIngresosGlobal]").Comando314_Click

'''''''''''''''PROCESAMIENTOS VS INGRESOS''''''''''''''''''''''''''''''

DoCmd.OpenForm "Control Procesamiento / Ingresos Global"
[Forms]![Control Procesamiento / Ingresos Global].asemana = Me.asemaprod
Call Forms("[Control Procesamiento / Ingresos Global]").Comando206_Click

'''''''''''''''PROCESAMIENTO VS PORCIONES''''''''''''''''''''''''''''''

DoCmd.OpenForm "ProcesamientovsPorcionamiento"
[Forms]![ProcesamientovsPorcionamiento].asemana = Me.asemaprod
Call Forms("[ProcesamientovsPorcionamiento]").Comando295_Click

'''''''''''''''CONTROL PORCIONES''''''''''''''''''''''''''''''

DoCmd.OpenForm "ControlInventarioPorciones"
[Forms]![ControlInventarioPorciones].semanaac = Me.asemaprod
Call Forms("[ControlInventarioPorciones]").Comando196_Click

'''''''''''''''CONTROL NO PORCIONABLES''''''''''''''''''''''''''''''

Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantix As Integer
Dim grupi As String
Dim ax As Integer

miSQL = "SELECT DBIngredientes.Tipo, DBIngredientes.TIPO1" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.Tipo, DBIngredientes.TIPO1" & _
" HAVING (DBIngredientes.TIPO1)='VARIABLES'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantix = rst.RecordCount
rst.MoveFirst

For ax = 1 To cantix
    grupi = rst("Tipo")
    DoCmd.OpenForm "Control Semanal Main"
    [Forms]![Control Semanal Main].semanaac = Me.asemaprod
    [Forms]![Control Semanal Main].Tipo = grupi
    Call Forms("[Control Semanal Main]").Comando157_Click
    rst.MoveNext
Next ax

rst.Close

'''''''''''''''''''''''''''FIN'''''''''''''''''''''''''

Dim beepc As Integer
beepc = 0
Do Until beepc = 10
    beepc = beepc + 1
    DoCmd.Beep
    Sleep 1000
Loop
MsgBox "Reportes Gurdados Correctamente"
End Sub

Private Sub Comando2554_Click()
DoCmd.OpenForm "ControlProductosGenerales"
End Sub













Public Sub Comando274_Click()
Dim Direccion As String
Dim Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.mhasta
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.mhasta & " - 11.Ventas Ultimas " & Me.mhasta - Me.mdesde + 1 & " Semanas.jpg"
DoCmd.OpenForm "Dashboard_Ventas_Totales", acFormPivotChart, , "[semana]>=" & Me.mdesde & " AND [semana]<=" & Me.mhasta
[Forms]![Dashboard_Ventas_Totales].Form.Requery
[Forms]![Dashboard_Ventas_Totales].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_Totales"

Me.Requery
Call EnviarImagenTelegram(Direccion & "\", Me.mhasta & " - 11.Ventas Ultimas " & Me.mhasta - Me.mdesde + 1 & " Semanas.jpg", "Ventas de ultimas " & Me.mhasta - Me.mdesde + 1 & " semanas " & nombrelocal(), grupotgerencia)
End Sub

Private Sub Comando1449_Click()

Dim Direccionano, Direccionmes As String
Dim archivo As String

Direccionano = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anhasta
Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Mensual\" & Me.anhasta & "\" & Me.mehasta
If Dir(Direccionano, vbDirectory) = "" Then
    MkDir Direccionano
End If
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

''''''''''''''''''''''''''''''''''VENTAS MENSUALES
archivo = Direccionmes & "\P" & codigoLocal() & " " & Me.mehasta & "_" & Me.anhasta & " - Ventas Mensuales.jpg"
DoCmd.OpenForm "Dashboard_Ventas_Mes", acFormPivotChart, , "[periodo]>=" & Me.andesde * 100 + Me.medesde & " AND [periodo]<=" & Me.anhasta * 100 + Me.mehasta
[Forms]![Dashboard_Ventas_Mes].Form.Requery
[Forms]![Dashboard_Ventas_Mes].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_Mes"

''''''''''''''''''''''''''''''''''VENTAS DIARIAS PROMEDIO
archivo = Direccionmes & "\P" & codigoLocal() & " " & Me.mehasta & "_" & Me.anhasta & " - Ventas Diarias Promedio.jpg"
DoCmd.OpenForm "Dashboard_Ventas_DiaPromedio", acFormPivotChart, , "[periodo]>=" & Me.andesde * 100 + Me.medesde & " AND [periodo]<=" & Me.anhasta * 100 + Me.mehasta
[Forms]![Dashboard_Ventas_DiaPromedio].Form.Requery
[Forms]![Dashboard_Ventas_DiaPromedio].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_DiaPromedio"

''''''''''''''''''''''''''''''''''VENTAS MENSUAL GRUPO
'archivo = Direccionmes & "\P" & codigolocal() & " " & Me.mehasta & "_" & Me.anhasta & " - Ventas Mensuales Grupos (Cantidad).jpg"
'DoCmd.OpenForm "Dashboard_Ventas_MesGrupos", acFormPivotChart, , "[periodo]>=" & Me.andesde * 100 + Me.medesde & " AND [periodo]<=" & Me.anhasta * 100 + Me.mehasta
'[Forms]![Dashboard_Ventas_MesGrupos].Form.Requery
'[Forms]![Dashboard_Ventas_MesGrupos].ChartSpace.exportpicture archivo
'DoCmd.Close acForm, "Dashboard_Ventas_MesGrupos"

''''''''''''''''''''''''''''''''''VENTAS PROMOCION CANTIDAD
'archivo = Direccionmes & "\P" & codigolocal() & " " & Me.mehasta & "_" & Me.anhasta & " - Ventas Mensuales Promocion (Cantidad).jpg"
'DoCmd.OpenForm "Dashboard_Cantidad_Promociones", acFormPivotChart, , "[periodo]>=" & Me.andesde * 100 + Me.medesde & " AND [periodo]<=" & Me.anhasta * 100 + Me.mehasta
'[Forms]![Dashboard_Cantidad_Promociones].Form.Requery
'[Forms]![Dashboard_Cantidad_Promociones].ChartSpace.exportpicture archivo
'DoCmd.Close acForm, "Dashboard_Cantidad_Promociones"
End Sub

Private Sub Comando1554_Click()
DoCmd.OpenForm "AgregarReceta"
End Sub

Private Sub Comando1603_Click()
DoCmd.OpenForm "Control Existencias Compra Venta"
End Sub

Private Sub Comando1627_Click()
DoCmd.OpenForm "AlertaInsumosAgotados"
End Sub

Private Sub Comando1630_Click()
Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.mhasta
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.mhasta & " - Volumen Insumos de " & Me.volsemhasta - Me.volsemdesde + 1 & " Semanas.jpg"
DoCmd.OpenForm "Dashboard_Consumos_Semanales", acFormPivotChart, , "[semana]>=" & Me.volsemdesde & " AND [semana]<=" & Me.volsemhasta
[Forms]![Dashboard_Consumos_Semanales].Form.Requery
[Forms]![Dashboard_Consumos_Semanales].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Consumos_Semanales"
End Sub

Private Sub Comando1669_Click()
DoCmd.OpenForm "Control de Ingredientes"
End Sub

Private Sub Comando1676_Click()
DoCmd.OpenForm "Indicadores Micro"
End Sub

Private Sub Comando1678_Click()
DoCmd.OpenForm "Indicadores Macro"
End Sub

Private Sub Comando1681_Click()
DoCmd.OpenForm "Relacion Afiliados"
End Sub

Private Sub Comando1683_Click()
DoCmd.OpenReport "VentasCliente", acViewNormal
End Sub

Private Sub Comando1909_Click()
DoCmd.OpenForm "Calculo Ventas Diarias Tipo Semana"
End Sub





Private Sub Comando1919_Click()
DoCmd.OpenForm "MenuPitayaProduccion0"
End Sub

Private Sub Comando1924_Click()
MsgBox numerosemana(Me.calfecha)
End Sub

Private Sub Comando1943_Click()

Call eliminartablasmain
Call importartablasmain

Call eliminartablascentral(codigoLocal())
Call importartablascentral(codigoLocal())

If APIDisponible() Then
    Call importartablasweb
Else
    Call importartablaespecifica("RRHH", "Operarios", "Operarios", 3)
    Call importartablaespecifica("RRHH", "AsignacionNivelesCargos", "AsignacionNivelesCargos", 3)
    Call importartablaespecifica("RRHH", "NivelesCargos", "NivelesCargos", 3)
End If

MsgBox "Tablas Main Actualizadas"
End Sub

















Private Sub Comando1990_Click()
DoCmd.OpenForm "HistorialdeProductosVenta"
End Sub





Private Sub Comando2766_Click()
Call EnviarMensajeWhatsapp("Buenos Dias", Me.wspme)
End Sub



Private Sub Comando2799_Click()
Call eliminartablasmain

Call vinculartablasmain

MsgBox "Tablas Main Actualizadas"
End Sub


Private Sub Comando2806_Click()
'On Error GoTo Nulo
Dim cont As Integer

For cont = Me.clubdesde To Me.clubhasta
    DoCmd.OpenReport "StickerClientesClub", acViewReport
    [Reports]![StickerClientesClub]![mcodigo] = cont
    DoCmd.SelectObject acReport, "StickerClientesClub"
    DoCmd.PrintOut acSelection, 1, 1
    DoCmd.Close acReport, "StickerClientesClub"
Next cont

Exit Sub
Nulo:
MsgBox "No se imprimio, se produjo un error"
End Sub

Private Sub Comando2820_Click()
DoCmd.OpenForm "Relacion de Productos Venta"
End Sub



Private Sub Comando2837_Click()
DoCmd.OpenForm "ResumenPreIngresosPitaya"
[Forms]![ResumenPreIngresosPitaya].adestino = "Pitaya " & codigoLocal()
[Forms]![ResumenPreIngresosPitaya].adestino.Locked = True
[Forms]![ResumenPreIngresosPitaya]![aencabezado].Caption = "RESUMEN DE INGRESOS DE LA SEMANA"
End Sub

Private Sub Comando2840_Click()
Dim Direccion As String
Dim Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.mghasta
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.mghasta & " - Ventas Ultimas Consolidado " & Me.mghasta - Me.mgdesde + 1 & " Semanas.jpg"
DoCmd.OpenForm "Dashboard_Ventas_Sucursales", acFormPivotChart, , "[semana]>=" & Me.mgdesde & " AND [semana]<=" & Me.mghasta

[Forms]![Dashboard_Ventas_Sucursales].Form.Requery
[Forms]![Dashboard_Ventas_Sucursales].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_Sucursales"

Me.Requery
Call EnviarImagenTelegram(Direccion & "\", Me.mghasta & " - Ventas Ultimas Consolidado " & Me.mghasta - Me.mgdesde + 1 & " Semanas.jpg", "Ventas consolidadas de ultimas " & Me.mghasta - Me.mgdesde + 1 & " semanas", grupotgerencia())
End Sub

Public Sub Comando2879_Click()
On Error GoTo Error

Dim Direccionano As String
Dim Direccionmes As String

Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Ventas"
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

If codigoLocal() = 15 Then
    DoCmd.OutputTo acOutputQuery, "ResumenVentasMesExcelAtencionAlCliente", acFormatXLSX, Direccionmes & "\" & nombrelocalglobal(codigoLocal()) & " - Historial Ventas Desde " & Format(Me.excelventasdesde, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.excelventashasta, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False
Else
    DoCmd.OutputTo acOutputQuery, "ResumenVentasMesExcel", acFormatXLSX, Direccionmes & "\" & nombrelocalglobal(codigoLocal()) & " - Historial Ventas Desde " & Format(Me.excelventasdesde, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.excelventashasta, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False
End If

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"
End Sub

Public Sub Comando2880_Click()
On Error GoTo Error

Dim Direccionano As String
Dim Direccionmes As String

Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Compras"
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If


DoCmd.OutputTo acOutputQuery, "HistorialComprasDatos", acFormatXLSX, Direccionmes & "\" & nombrelocalglobal(codigoLocal()) & " - Historial Compras Desde " & Format(Me.excelcomprasdesde, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.excelcomprashasta, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente"
End Sub

Private Sub Comando2881_Click()
DoCmd.OpenForm "HistorialComprasGobalDetalle", acNormal
[Forms]![HistorialComprasGobalDetalle].Form.Filter = "[semana]=" & Me.semasucur & " AND [local]<>0"
[Forms]![HistorialComprasGobalDetalle].Form.FilterOn = True
[Forms]![HistorialComprasGobalDetalle].Form.Requery
[Forms]![HistorialComprasGobalDetalle].[Pagado].width = 0
[Forms]![HistorialComprasGobalDetalle].[semana].width = 0
[Forms]![HistorialComprasGobalDetalle].[CodCotizacion].width = 0

[Forms]![HistorialComprasGobalDetalle].Form.Printer.Orientation = acPRORLandscape

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\Pitaya 0\" & Me.semasucur
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\Pitaya 0"
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.semasucur & " - Historial Facturas Registradas Sucursales.pdf"
DoCmd.OutputTo acOutputForm, "HistorialComprasGobalDetalle", acFormatPDF, archivo, False

DoCmd.Close acForm, "HistorialComprasGobalDetalle", acSaveNo
End Sub

Public Sub Comando2882_Click()
On Error GoTo Error

Dim Direccion, Direccion2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Inventarios"
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

DoCmd.OutputTo acOutputQuery, "HistorialInventariosDatos", acFormatXLSX, Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - " & Me.invseman & " - Inventario Final.xlsx", False

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente"
End Sub

Public Sub Comando2946_Click()
On Error GoTo Error

Dim Direccion, Direccion2 As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\PreIngresos"
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

DoCmd.OutputTo acOutputQuery, "HistorialPreingresosDatos", acFormatXLSX, Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - " & Me.preingseman & " Pregingresos Semanales.xlsx", False

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente"
End Sub

Private Sub Comando2963_Click()
DoCmd.OpenForm "Relacion de Productos Venta x Insumo"
End Sub

Private Sub Comando2966_Click()
Dim indi As String

Select Case CurrentProject.Name
    Case "Pitaya_System.accdb" 'archivo base de sistema
        'NO VINCULA NADA AL MIXED
        MsgBox "Sistema Raiz no permite vincular tabla a db mixed"
    Case Else
        indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
        Call vinculartablasLocalMixed(CInt(indi))
        MsgBox "Importacion Vinculada de Tablas Finalizada"
End Select



End Sub

Private Sub Comando3014_Click()
Dim Direccion As String
Dim Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.mshasta
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.mshasta & " - Ventas Ultimas Consolidado Suma " & Me.mshasta - Me.msdesde + 1 & " Semanas.jpg"
DoCmd.OpenForm "Dashboard_Ventas_Sucursales_Suma", acFormPivotChart, , "[semana]>=" & Me.msdesde & " AND [semana]<=" & Me.mshasta
[Forms]![Dashboard_Ventas_Sucursales_Suma].Form.Requery
[Forms]![Dashboard_Ventas_Sucursales_Suma].ChartSpace.exportpicture archivo
DoCmd.Close acForm, "Dashboard_Ventas_Sucursales_Suma"

Me.Requery
Call EnviarImagenTelegram(Direccion & "\", Me.mshasta & " - Ventas Ultimas Consolidado Suma " & Me.mshasta - Me.msdesde + 1 & " Semanas.jpg", "Ventas suma de ultimas " & Me.mshasta - Me.msdesde + 1 & " semanas", grupotgerencia())

End Sub

Private Sub Comando3033_Click()
DoCmd.OpenForm "EleccionSucursalVigente"
[Forms]![EleccionSucursalVigente].tipocarga = 1
End Sub

Private Sub Comando3052_Click()
DoCmd.OpenForm "HistorialComprasGobalDetalle", acNormal
[Forms]![HistorialComprasGobalDetalle].Form.Filter = "[Fecha]>=#" & Me.fechacompragdesde & "# AND [Fecha]<=#" & Me.fechacompraghasta & "# AND [local]<>0"
[Forms]![HistorialComprasGobalDetalle].Form.FilterOn = True
[Forms]![HistorialComprasGobalDetalle].Form.Requery
[Forms]![HistorialComprasGobalDetalle].[Pagado].width = 0
[Forms]![HistorialComprasGobalDetalle].[semana].width = 0
[Forms]![HistorialComprasGobalDetalle].[CodCotizacion].width = 0
[Forms]![HistorialComprasGobalDetalle].Form.Printer.Orientation = acPRORLandscape

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\Pitaya 0\" & numerosemana(Me.fechacompragdesde)
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\Pitaya 0"
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & numerosemana(Me.fechacompragdesde) & " - Historial Diario de Facturas Registradas Sucursales Desde " & Format(Me.fechacompragdesde, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.fechacompraghasta, "dd"" de ""mmmm"" de ""yyyy") & ".pdf"
DoCmd.OutputTo acOutputForm, "HistorialComprasGobalDetalle", acFormatPDF, archivo, False

MsgBox "Guardado en la carpeta de semana " & numerosemana(Me.fechacompragdesde)
DoCmd.Close acForm, "HistorialComprasGobalDetalle", acSaveNo
End Sub

Private Sub Comando3064_Click()
DoCmd.OpenForm "HistoricoConsumoPorcionesDescargado"
End Sub

Private Sub Comando3068_Click()
DoCmd.OpenForm "HistoricoConsumoNoPorcionesDescargado"
End Sub

Private Sub Comando3071_Click()
DoCmd.OpenForm "HistoricoConsumoMostradorDescargado"
End Sub

Private Sub Comando3100_Click()


On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.Nombre, DatosSistema.IngresoInsumos, DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0) And ((DatosSistema.CodSistema)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    'Porciones
    DoCmd.OpenForm "HistoricoConsumoPorcionesSucursales"
    [Forms]![HistoricoConsumoPorcionesSucursales].asucursal = rst("CodSistema")
    Call Forms("[HistoricoConsumoPorcionesSucursales]").Comando170_Click
    
    'No Porciones
    DoCmd.OpenForm "HistoricoConsumoNoPorcionesSucursales"
    [Forms]![HistoricoConsumoNoPorcionesSucursales].alocali = rst("CodSistema")
    Call Forms("[HistoricoConsumoNoPorcionesSucursales]").Comando170_Click
    
    'Mostrador
    DoCmd.OpenForm "HistoricoConsumoMostradorSucursales"
    [Forms]![HistoricoConsumoMostradorSucursales].alocali = rst("CodSistema")
    Call Forms("[HistoricoConsumoMostradorSucursales]").Comando170_Click
    
    rst.MoveNext
Loop
rst.Close
Exit Sub

Nulo:
MsgBox "Error en la generacion de historicos"
End Sub

Private Sub Comando3113_Click()
Call Comando3014_Click
Call Comando2840_Click
End Sub

Private Sub Comando3229_Click()
DoCmd.OpenForm "RegistroEstandarInsumosFijosSemana"
End Sub



Private Sub Comando3231_Click()
On Error GoTo Error
Dim ax As Integer

'Vacia la tabla existente en la base de datos local
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM [tClientesClub]"
DoCmd.SetWarnings True

For ax = 1 To cantidadsucursalesexistentes()
    Call importartablaespecifica("Pitaya" & ax & "_DB", "ClientesClub", "tClientesClub", 2)
Next ax

Dim direccionclub As String
direccionclub = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Club Pitaya"
If Dir(direccionclub, vbDirectory) = "" Then
    MkDir direccionclub
End If

DoCmd.OutputTo acOutputTable, "tClientesClub", acFormatXLSX, direccionclub & "\Lista de Club Pitaya actualizado a " & Format(Now(), "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False
MsgBox "Datos Exportados en la carpeta de Club Pitaya"
Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"

End Sub

Private Sub Comando3243_Click()
Dim nombrearchi As String
nombrearchi = CurrentProject.Name

On Error GoTo Nulo
Dim ax As Integer
For ax = 1 To cantidadsucursalesexistentes()

    If nombrearchi Like "Modulo*" Or nombrearchi = "Pitaya_System.accdb" Then

        Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "ClientesClub", "ClientesClub" & ax)
        Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "NotaDePedido", "NotaDePedido" & ax)
        Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "SubPedido", "SubPedido" & ax)
    Else
        Dim indi As String
        indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
        If indi <> ax Then
            Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "ClientesClub", "ClientesClub" & ax)
            Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "NotaDePedido", "NotaDePedido" & ax)
            Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "SubPedido", "SubPedido" & ax)
        End If
    End If
    
Next ax
MsgBox "FIn"
Exit Sub

Nulo:
MsgBox "Error al importar datos de ventas club al importar base de datos " & ax
End Sub

Private Sub Comando3247_Click()
Dim semanab As Integer
semanab = InputBox("Ingresar numero de semana, semana actual: " & numerosemana(Date))

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Main_DB2.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim dbexterna As DAO.Database
Dim rst As DAO.Recordset
Dim miSQL As String
Set dbexterna = DBEngine.OpenDatabase(h)
miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE (StatusSucursales.Sucursal<>0 AND StatusSucursales.Activo<>0)"
Set rst = dbexterna.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF

    If CurrentProject.Name = "Pitaya_System.accdb" Then 'archivo base de sistema
            Call actualizartablas(1, rst("CodLocal"))
    Else
        MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
        Exit Sub
    End If


    DoCmd.OpenForm "Control Insumos Importantes"
    [Forms]![Control Insumos Importantes]![semana1] = semanab
    Call Forms("[Control Insumos Importantes]").Comando562_Click
    
    'DoCmd.OpenForm "Control Porcionamiento"
    '[Forms]![Control Porcionamiento]![asemana] = semanab
    'Call Forms("[Control Porcionamiento]").Comando133_Click
    
    DoCmd.OpenForm "Control Existencias Compra Venta"
    [Forms]![Control Existencias Compra Venta]![asemana] = semanab
    Call Forms("[Control Existencias Compra Venta]").Comando352_Click
    
    'DoCmd.OpenForm "ControlProductosGenerales"
    '[Forms]![ControlProductosGenerales]![asemana] = semanab
    'Call Forms("[ControlProductosGenerales]").Comando314_Click
    
    rst.MoveNext
Loop
dbexterna.Close
rst.Close
Set fs = Nothing
Kill (h)

MsgBox "PDF Generados correctamente"
Exit Sub

Nulo:
rst.Close
Set fs = Nothing
Kill (h)

MsgBox "No se han generado los PDF correctamente"
End Sub

Private Sub Comando3285_Click()
On Error GoTo Nulo
'Habilitar Referencia Microsoft XML v6.0
'Descatvar privacidad de grupo de bot en telegram
Dim url As String
Const RunAsync As Boolean = True
Const ProcessComplete As Integer = 4

Dim request As MSXML2.XMLHTTP60
Set request = New MSXML2.XMLHTTP60

Dim response As String
Dim token As String

token = DLookup("[IdBotTe]", "DatosSistema")

url = "https://api.telegram.org/bot" & token & "/getUpdates"
url = url & "?cb=" & Timer() * 100
With request
    .Open "GET", url, RunAsync
    .setRequestHeader "Content-Type", "application/json"
    .send
    Do While request.ReadyState <> ProcessComplete
        DoEvents
    Loop
    response = .responseText
End With
Set request = Nothing

MsgBox response

Dim divisionactu() As String
Dim ccuenta As Integer
Dim ultimaactu As Long

divisionactu = Split(response, "update_id" & Chr(34) & ":")
ccuenta = UBound(divisionactu) - LBound(divisionactu) + 1
ultimaactu = Split(Split(response, "update_id" & Chr(34) & ":")(ccuenta - 1), ",")(0)
Call LimpiarHistorialTelegram(ultimaactu)

Exit Sub

Nulo:
End Sub

Private Sub Comando3288_Click()
DoCmd.OpenForm "PruebaVelocidadFormulas"
End Sub

Private Sub Comando3291_Click()


MsgBox Application.Printer.DeviceName
End Sub

Private Sub Comando3318_Click()
DoCmd.OpenForm "HistorialRendimientoProduccion"
End Sub

Private Sub Comando3320_Click()
Dim semanab As Integer
semanab = InputBox("Ingresar numero de semana, semana actual: " & numerosemana(Date))

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DatosSistema.CodSistema, DatosSistema.Activo" & _
" FROM DatosSistema IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" WHERE (((DatosSistema.Activo)<>0) And ((DatosSistema.CodSistema)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF

    If CurrentProject.Name = "Pitaya_System.accdb" Or CurrentProject.Name = "Pitaya_System - Copy.accdb" Or CurrentProject.Name = "Pitaya_System - Copia.accdb" Then 'archivo base de sistema
            Call actualizartablas(1, rst("CodSistema"))
    Else
        MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
        Exit Sub
    End If


    DoCmd.OpenForm "HistorialRendimientoProduccion"
    [Forms]![HistorialRendimientoProduccion]![asemana] = semanab
    
    Call Forms("[HistorialRendimientoProduccion]").Comando206_Click
    rst.MoveNext
Loop
rst.Close
Exit Sub

Nulo:
MsgBox "PDF Generados"
End Sub

Private Sub Comando3322_Click()
'StickerEquiposCodigo
'On Error GoTo Nulo

Dim cont As Integer

For cont = Me.acodigoequipoini To Me.acodigoequipofin
    DoCmd.OpenReport "StickerEquiposCodigo", acViewReport
    [Reports]![StickerEquiposCodigo]![acodigo] = Me.acodigoequiposigla & cont
    [Reports]![StickerEquiposCodigo]![aequipo] = Me.aequipoc
    DoCmd.SelectObject acReport, "StickerEquiposCodigo"
    DoCmd.PrintOut acSelection, 1, 1
    DoCmd.Close acReport, "StickerEquiposCodigo"
Next cont


Exit Sub
Nulo:
MsgBox "No se imprimio, se produjo un error"
End Sub

Public Sub Comando3390_Click()

On Error GoTo Error

'Descargar listas despacho actualizadaso
Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)

Dim Direccionano As String
Dim Direccionmes As String

Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Despachos"
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

DoCmd.OutputTo acOutputQuery, "ResumenDespachoFechaExcel", acFormatXLSX, Direccionmes & "\" & nombrelocalglobal(codigoLocal()) & " - Historial Despacho Desde " & Format(Me.despainicio, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.despafinal, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"
End Sub

Private Sub Comando3410_Click()
DoCmd.OpenForm "Calculo Consumo NoPorcion Semana"
End Sub

Private Sub Comando3418_Click()
DoCmd.OpenForm "HistorialVentasPedido"
End Sub

Private Sub Comando3429_Click()
If CurrentProject.Name = "Pitaya_System.accdb" Or CurrentProject.Name = "Pitaya_System - Copy.accdb" Or CurrentProject.Name = "Pitaya_System - Copia.accdb" Then 'archivo base de sistema
        Call actualizartablas(1, 0)
Else
    MsgBox "Tiene que estar con el sistema raiz para descargar datos de la central"
    Exit Sub
End If


On Error GoTo Error
Dim Direccionmes As String

Direccionmes = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Produccion"
If Dir(Direccionmes, vbDirectory) = "" Then
    MkDir Direccionmes
End If

DoCmd.OutputTo acOutputQuery, "ResumenPorcionamientoFechaExcel", acFormatXLSX, Direccionmes & "\" & nombrelocalglobal(codigoLocal()) & " - Historial Porcionamiento Desde " & Format(Me.dproduc, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.hprodu, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False
DoCmd.OutputTo acOutputQuery, "ResumenProcesamientoFechaExcel", acFormatXLSX, Direccionmes & "\" & nombrelocalglobal(codigoLocal()) & " - Historial Procesamiento Desde " & Format(Me.dproduc, "dd"" de ""mmmm"" de ""yyyy") & " Hasta " & Format(Me.hprodu, "dd"" de ""mmmm"" de ""yyyy") & ".xlsx", False
MsgBox "Datos Exportados en la carpeta de Produccion"
Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"
End Sub

Public Sub Comando3461_Click()
On Error GoTo Error
Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Datos Descargados de Sistema\Balance Semanal"
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

DoCmd.OutputTo acOutputQuery, "ExcelControlinsumosImportantes", acFormatXLSX, Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - " & Me.semanacontrolimpor & " - Control de Insumos Importantes.xlsx", False
DoCmd.OutputTo acOutputQuery, "ExcelControlinsumosMostrador", acFormatXLSX, Direccion & "\" & nombrelocalglobal(codigoLocal()) & " - " & Me.semanacontrolimpor & " - Control de Insumos Mostrador.xlsx", False

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"
End Sub



Public Sub Comando3498_Click()
On Error GoTo Error
Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.semanahorario
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

DoCmd.OutputTo acOutputQuery, "ExcelHorarioMarcacionesSemana", acFormatXLSX, Direccion & "\" & "Horario Marcado Semana.xlsx", False

Exit Sub
Error:
MsgBox "No se exporto el excel correctamente, vuelva a descargar excel"
End Sub

Private Sub Comando3510_Click()
Call funcionrotacionsucursalesactivas("Comando3498_Click")
End Sub



Private Sub Comando3814_Click()
Call funcionrotacionsucursalesactivas("Comando2879_Click")

End Sub

Private Sub Comando3816_Click()
Call actualizartablas(1, 0)
Call Comando2880_Click
Call funcionrotacionsucursalesactivas("Comando2880_Click")
End Sub

Private Sub Comando3818_Click()
Call funcionrotacionsucursalesactivas("Comando3390_Click")
End Sub

Private Sub Comando3496_Click()
Call funcionrotacionsucursalesactivas("Comando3461_Click")
End Sub

Private Sub Comando3820_Click()
Call funcionrotacionsucursalesactivas("Comando2882_Click")
End Sub

Private Sub Comando3822_Click()
Call funcionrotacionsucursalesactivas("Comando2946_Click")
End Sub

Private Sub Comando3825_Click()
DoCmd.OpenForm "VerificarMarcacionesSemana"
End Sub



Private Sub Comando3829_Click()
On Error GoTo Nulo
Dim busca As String
Dim indi As String
Call eliminararchivosllavelectura
busca = CurrentProject.Name

If busca Like "Pitaya*" Then 'empieza con pitaya entonces si puede er sucursal o raiz\
    busca = Left(Right(busca, Len(busca) - 6), 1)
    If IsNumeric(busca) Then
        indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
        Call actualizartablas(3, CInt(indi))
    End If
Else
    MsgBox "Tiene que ser un archivo de pitaya local para cargar datos de local en edicion"
End If

Exit Sub
Nulo:
MsgBox "Problemas al cargar datos"

End Sub

Private Sub Comando3849_Click()
DoCmd.OpenForm "EleccionSucursalVigente"
[Forms]![EleccionSucursalVigente].tipocarga = 1
End Sub

Private Sub Comando3852_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "GestionSistema"
End Sub





Private Sub Comando3876_Click()
'Call SincronizarCedulasLocales
Call enviarnotificacionclienteusopuntos(1402, 20, "Altamira")
End Sub

Private Sub Comando4107_Click()
DoCmd.CopyObject , "NotaDePedidoRevision", acTable, "NotaDePedido"
DoCmd.OpenForm "HistorialVentasPedidoxSucursal"
[Forms]![HistorialVentasPedidoxSucursal].Form.Filter = "[Fecha]>= #" & Date & "# And [Fecha]<= #" & Date & "#"
[Forms]![HistorialVentasPedidoxSucursal].Form.FilterOn = True
End Sub


Private Sub Comando4130_Click()
Call ReiniciarGoogleDrive
End Sub

Private Sub Comando4150_Click()
DoCmd.OpenForm "ReporteDeliveryPitaya"
End Sub

Private Sub Comando4162_Click()
DoCmd.OpenForm "Control Semanal Main"
End Sub

Private Sub Comando4163_Click()
DoCmd.OpenForm "ControlInventarioPorciones"
End Sub

Private Sub Comando4185_Click()
DoCmd.OpenForm "ReporteVentasPedidoSucursales"
End Sub

Private Sub Comando4191_Click()
DoCmd.OpenForm "ReporteVentasPedidoSucursales"
End Sub

Private Sub Comando4207_Click()
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\CopiaOtro" & NombrePC() & NombreSistema() & "Main_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim db As Database
Dim td As TableDef
Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs 'recorrer nombres de tablas main copia
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        Call importartablaespecificaAMixed("Main_DB", td.Name, td.Name)

    End If
Next

db.Close
Set fs = Nothing
Kill (h)

MsgBox "Tablas Main actualizados en archivo Mixed"
End Sub

Private Sub Comando4209_Click()
DoCmd.OpenForm "StatusDriveSucursales"
End Sub
    
Private Sub Comando4345_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?") = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("SubReceta", "SubReceta", "", 2)
End Sub



Private Sub Comando4424_Click()
Dim fechaactual As Date
If MsgBox("Desea descargar archivo csv segun fechas ingresadas, GENERAR UN ARCHIVO POR CADA FECHA PARA AUTOMATICO", vbYesNo, "Exportacion de datos") <> vbYes Then
    Exit Sub
End If

If Me.vdesdehost > Me.vhastahost Then
    MsgBox "Fechas desde es mayor de la fecha ahsta"
    Exit Sub
End If

If IsNull(Me.vdesdehost) Or IsNull(Me.vhastahost) Then
    MsgBox "Debe ingresar ambas fechas"
    Exit Sub
End If

' Recorrer directamente desde los controles del formulario
For fechaactual = Me.vdesdehost To Me.vhastahost
    Call ActualizarArchivoCSVVentasFecha(fechaactual)
Next fechaactual

End Sub



Private Sub Comando4418_Click()
Dim fechaactual As Date
If MsgBox("Desea descargar archivo csv segun fechas ingresadas, UN SOLO ARCHIVO CON TODAS LAS FECHAS", vbYesNo, "Exportacion de datos") <> vbYes Then
    Exit Sub
End If

If Me.vdesdehost > Me.vhastahost Then
    MsgBox "Fechas desde es mayor de la fecha ahsta"
    Exit Sub
End If

If IsNull(Me.vdesdehost) Or IsNull(Me.vhastahost) Then
    MsgBox "Debe ingresar ambas fechas"
    Exit Sub
End If

'Ruta de donde se guardara el csv en pendiente de subida
Dim ruta As String
Dim condic As String
ruta = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Web\Descargas Manuales de Access\" & codigoLocal() & " desde " & cfechasqlfecha(Me.vdesdehost) & " hasta " & cfechasqlfecha(Me.vhastahost) & ".csv"
condic = "WHERE Fecha BETWEEN '" & cfechasqlfecha(Me.vdesdehost) & "' AND '" & cfechasqlfecha(Me.vhastahost) & "'"
Call ExportarCSVConsultaCondicional("ResumenVentasMesExcelHostingerCSV", condic, ruta)

End Sub

Private Sub Comando4419_Click()

If MsgBox("Desea descargar archivo csv para subida manual desde 2016 al 2024 por año", vbYesNo, "Exportacion de datos") <> vbYes Then
    Exit Sub
End If

'2016
Me.vdesdehost = #1/1/2016#
Me.vhastahost = #12/31/2016#
Call Comando4418_Click

'2017
Me.vdesdehost = #1/1/2017#
Me.vhastahost = #12/31/2017#
Call Comando4418_Click

'2018
Me.vdesdehost = #1/1/2018#
Me.vhastahost = #12/31/2018#
Call Comando4418_Click

'2019
Me.vdesdehost = #1/1/2019#
Me.vhastahost = #12/31/2019#
Call Comando4418_Click

'2020
Me.vdesdehost = #1/1/2020#
Me.vhastahost = #12/31/2020#
Call Comando4418_Click

'2021
Me.vdesdehost = #1/1/2021#
Me.vhastahost = #12/31/2021#
Call Comando4418_Click

'2022
Me.vdesdehost = #1/1/2022#
Me.vhastahost = #12/31/2022#
Call Comando4418_Click

'2023
Me.vdesdehost = #1/1/2023#
Me.vhastahost = #12/31/2023#
Call Comando4418_Click

'2024
Me.vdesdehost = #1/1/2024#
Me.vhastahost = #12/31/2024#
Call Comando4418_Click

MsgBox "listo desde el2016 al 2024"
End Sub





Private Sub Comando4323_Click()
DoCmd.OpenForm "HistorialdeProductosVenta"
End Sub

Private Sub Comando4340_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("DBBatidos", "DBBatidos", "", 2)
End Sub

Private Sub Comando4506_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("DBIngredientes", "DBIngredientes", "", 2)
End Sub

Private Sub Comando4508_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("Cotizaciones", "Cotizaciones", "", 2)
End Sub

Private Sub Comando4512_Click()
Dim semanab As Integer
semanab = InputBox("Ingresar numero de semana, semana actual: " & numerosemana(Date))

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Main_DB2.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim dbexterna As DAO.Database
Dim rst As DAO.Recordset
Dim miSQL As String
Set dbexterna = DBEngine.OpenDatabase(h)
miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE (StatusSucursales.Sucursal<>0 AND StatusSucursales.Activo<>0)"
Set rst = dbexterna.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF

    If CurrentProject.Name = "Pitaya_System.accdb" Then 'archivo base de sistema
            Call actualizartablas(1, rst("CodLocal"))
    Else
        MsgBox "Tiene que estar con el sistema raiz para revisar otras sucursales"
        Exit Sub
    End If


    DoCmd.OpenForm "Control Porcionamiento"
    [Forms]![Control Porcionamiento]![asemana] = semanab
    Call Forms("[Control Porcionamiento]").Comando133_Click
    
    DoCmd.OpenForm "ControlProductosGenerales"
    [Forms]![ControlProductosGenerales]![asemana] = semanab
    Call Forms("[ControlProductosGenerales]").Comando314_Click
    
    rst.MoveNext
Loop
dbexterna.Close
rst.Close
Set fs = Nothing
Kill (h)

MsgBox "PDF Generados correctamente"
Exit Sub

Nulo:
rst.Close
Set fs = Nothing
Kill (h)

MsgBox "No se han generado los PDF correctamente"
End Sub

Private Sub Comando4521_Click()
DoCmd.OpenForm "PanelDescargaExcelConsumoDirecto"
End Sub

Private Sub Comando4531_Click()
Call SincronizarDatosClientesLocales
End Sub

Private Sub Comando4533_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("DBPromociones", "promociones_access_csv", "", 2)
End Sub



Private Sub Comando4545_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("MezclaPorciones", "MezclaPorcionesAccess", "", 2)
End Sub

Private Sub Comando4561_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call SyncKardexTiendaMasivoCompleto
End Sub

Private Sub Comando4563_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call SyncKardexCentralContabilidadMasivo
End Sub

Private Sub Comando4565_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call SyncKardexCentralDespachoMasivo
End Sub

Private Sub Comando4569_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call SyncCierreDepositosTiendaMasivo
End Sub

Private Sub Comando4570_Click()
DoCmd.OpenForm "Control Consumo Diario"
End Sub

Private Sub Comando4606_Click()
If MsgBox("Desea subir los datos de exixstencias de un dia especifico?", vbYesNo, "CONFIRMACION") = vbYes Then
    Call SyncKardexTiendaDiaEspecifico(Me.diakardex)
    Call SyncVentasDia(Me.diakardex)
End If
End Sub

Private Sub Comando4626_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call SyncKardexProduccionMasivoCompleto
End Sub

Private Sub Comando4628_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?", vbYesNo) = vbNo Then
    Exit Sub
End If
Call SyncKardexCentralAlmacenMasivo
End Sub

Private Sub Comando4644_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?") = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("Proovedores", "msaccess_masivo_Depositos_Proveedores", "", 2)
End Sub

Private Sub Comando4646_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?") = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("CeCoCuentas", "CeCoCuentas", "", 2)
End Sub

Private Sub Comando4648_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?") = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("CeCoSubCuentas", "CeCoSubCuentas", "", 2)
End Sub

Private Sub Comando4651_Click()
If MsgBox("Se eliminaran los datos de la tabla en el host y se subiran los de la tabla local, desea continuar?") = vbNo Then
    Exit Sub
End If
Call exportarconsultacondicionalhostinger("CentroCostos", "CentroCostos", "", 2)
End Sub

Private Sub Comando477_Click()
DoCmd.OpenForm "Historial Ventas"
End Sub

Private Sub Comando716_Click()
DoCmd.OpenForm "Datos del SIstema"
DoCmd.MoveSize 17200, 500 'izquierda, arriba  '1440 twips = 96 pixel = 1 inch  1920 pixxeles son 23200 ubicacion
End Sub

Private Sub Comando730_Click()
DoCmd.OpenForm "Horario_Semana", acFormPivotTable, , "[Sucursal]=" & codigoLocal() & " AND [sema]=" & Me.horasema
[Forms]![Horario_Semana].Form.InsideWidth = 20000
[Forms]![Horario_Semana].Form.InsideHeight = 6800
End Sub

Private Sub Comando770_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "Pago Personal"
End Sub

Private Sub ImpresionAutomatica_Click()

'''''''''''''''PEDIDOS ANULADOS''''''''''''''''''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Relacion Pedidos Anulados"
    [Forms]![Relacion Pedidos Anulados].asemana = Me.bsemana
    Call Forms("[Relacion Pedidos Anulados]").Comando119_Click
    Call Forms("[Relacion Pedidos Anulados]").Comando437_Click
    DoCmd.Close acForm, "Relacion Pedidos Anulados"
End If
'''''''''''''''''''''''''''''''''PROMOCIONES'''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Resultados Promocion"
    [Forms]![Resultados Promocion].asemana = Me.bsemana
    Call Forms("[Resultados Promocion]").Comando278_Click
    DoCmd.Close acForm, "Resultados Promocion"
End If
''''''''''''''''''''''Pago Personal'''''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Pago Personal"
    [Forms]![Pago Personal].asemana = Me.bsemana
    Call Forms("[Pago Personal]").Comando262_Click
    DoCmd.Close acForm, "Pago Personal"
End If
'''''''''''''''''CU Cotizacion '''''''''''''''''''''
DoCmd.OpenForm "Calculo CU Cotizacion Semana NP"
[Forms]![Calculo CU Cotizacion Semana NP].asemana = Me.bsemana
Call Forms("[Calculo CU Cotizacion Semana NP]").Comando119_Click
Call Forms("[Calculo CU Cotizacion Semana NP]").Comando437_Click
DoCmd.Close acForm, "Calculo CU Cotizacion Semana NP"

''''''''''''''''''''''Conversion Cotizacion''''''''''''''
DoCmd.OpenForm "Calculo Conversion Cotizacion Semana"
[Forms]![Calculo Conversion Cotizacion Semana].asemana = Me.bsemana
Call Forms("[Calculo Conversion Cotizacion Semana]").Comando119_Click
Call Forms("[Calculo Conversion Cotizacion Semana]").Comando437_Click
DoCmd.Close acForm, "Calculo Conversion Cotizacion Semana"

''''''''''''''''''''''''''''''Pareto Semana'''''''
DoCmd.OpenForm "Calculo Pareto Semana"
[Forms]![Calculo Pareto Semana].asemana = Me.bsemana
Call Forms("[Calculo Pareto Semana]").Comando119_Click
Call Forms("[Calculo Pareto Semana]").Comando187_Click
DoCmd.Close acForm, "Calculo Pareto Semana"

''''''''''''''''''''''''Existencias Pareto''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Control Existencias Pareto"
    [Forms]![Control Existencias Pareto].asemana = Me.bsemana
    Call Forms("[Control Existencias Pareto]").Comando119_Click
    Call Forms("[Control Existencias Pareto]").Comando437_Click
    DoCmd.Close acForm, "Control Existencias Pareto"
End If
''''''''''''''''''''''''''''COMPRA VENTA''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Control Existencias Compra Venta"
    [Forms]![Control Existencias Compra Venta].asemana = Me.bsemana
    Call Forms("[Control Existencias Compra Venta]").Comando119_Click
    Call Forms("[Control Existencias Compra Venta]").Comando352_Click
    DoCmd.Close acForm, "Control Existencias Compra Venta"
End If
''''''''''''''''''REPORTE SEMANAL'''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Reporte Semanal"
    [Forms]![Reporte Semanal].asemana = Me.bsemana
    Call Forms("[Reporte Semanal]").Comando56_Click
    Call Forms("[Reporte Semanal]").Comando175_Click
    DoCmd.Close acForm, "Reporte Semanal"
End If

''''''''''''''''''CONTROL PORCIONAMI?NTO'''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    DoCmd.OpenForm "Control Porcionamiento"
    [Forms]![Control Porcionamiento].asemana = Me.bsemana
    Call Forms("[Control Porcionamiento]").Comando102_Click
    Call Forms("[Control Porcionamiento]").Comando133_Click
    DoCmd.Close acForm, "Control Porcionamiento"
End If

''''''''''''''''''VENTAS SEMANALES'''''''''''''''
If codigoLocal() <> 0 Then 'Sistema local
    Me.mdesde = Me.bsemana - 10
    Me.mhasta = Me.bsemana
    Call Me.Comando274_Click
End If

MsgBox "Reportes Guardados Correctamente"

End Sub

Private Sub Comando773_Click()
DoCmd.OpenForm "Calculo CU Cotizacion Semana NP"
End Sub

Private Sub Comando776_Click()
DoCmd.OpenForm "Calculo CU Ingrediente Semana"
End Sub

Private Sub Comando783_Click()
DoCmd.OpenForm "Calculo Conversion Cotizacion Semana"
End Sub

Private Sub Comando786_Click()
DoCmd.OpenForm "Calculo Consumo ingrediente Semana"
End Sub

Private Sub Comando789_Click()
DoCmd.OpenForm "Calculo Pareto Semana"
End Sub

Private Sub Comando792_Click()
DoCmd.OpenForm "Control Existencias Pareto"
End Sub

Private Sub Comando796_Click()
DoCmd.OpenForm "Control Existencias Compra Venta"
End Sub

Private Sub Comando800_Click()
DoCmd.OpenForm "Reporte Semanal"
End Sub

Private Sub Comando805_Click()
DoCmd.OpenForm "Calculo CU Cotizacion Semana P"
End Sub

Private Sub Comando841_Click()
DoCmd.OpenForm "Calculo Ventas Diarias Semana"
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error Resume Next
Form.Caption = "Batidos Pitaya"
Me.ShortcutMenu = False
Me.cfgrupo.RowSource = "SELECT Grupos.NombreGrupo, Grupos.CodGrupo FROM Grupos"
Me.Sucursal = ciudadsistema()


'1440 twips = 96 pixel = 1 inch
Me.InsideHeight = 11520
Me.InsideWidth = 16000
'MsgBox Me.WindowTop
'MsgBox Me.WindowLeft
Me.Move -645, -1100

End Sub

