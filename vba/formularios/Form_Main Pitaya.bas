' ==========================================================
' Modulo  : Form_Main Pitaya
' Tipo    : 100  |  Lineas: 586
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database

Private Sub Comando1711_Click()

DoCmd.OpenForm "Inicial"
End Sub



Private Sub Comando1713_Click()

On Error GoTo Nulo

If (Time - UltimoCierre()) * 24 * 60 < 60 Then 'da resultado minutos desde ultimo cierre a ahora, 1 hora despues del ultimo cierre no es permitido
    MsgBox "Existe un cierre de hace menos de una hora realizada, solicitar autorizacion para realizar cierre nuevamente"
    
    Exit Sub
End If

Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)

If existepreingresopendientefecha(Me.fechasistema) > 0 And Time() > #5:00:00 PM# Then
    
    MsgBox "Existen Preingresos que no han sido ingresados aun, corregir y dar ingreso a todos los ingresos pendientes de hoy para poder hacer cierre"

Else
    
    Call actualizarmontoguardadopedidosdia(Date)
    
    
    Dim pedidoscalculados As Double
    Dim pedidosguardados As Double
    
    pedidoscalculados = AcumuladoDia(Me.fechasistema)
    pedidosguardados = AcumuladoDiaGuardado(Me.fechasistema)
    
    If pedidoscalculados = pedidosguardados Then ' Datos cargados correctamente
        
        DoCmd.OpenForm "Cierre por Turno"
        [Forms]![Cierre por Turno]![aCodOperario] = [Forms]![Main Pitaya]![codigologin]
        [Forms]![Cierre por Turno]![aHoraInicial] = UltimoCierre()
        [Forms]![Cierre por Turno]![aHoraFinal] = Time()
        [Forms]![Cierre por Turno]![afecha] = Date
        [Forms]![Cierre por Turno]![montoperiodo] = pedidoscalculados
        
        [Forms]![Cierre por Turno]![aTotalPedidosYa] = 0
        [Forms]![Cierre por Turno]![aTotalPOS] = 0
        [Forms]![Cierre por Turno]![aTotalTransferencia] = 0
       
        [Forms]![Cierre por Turno]![possistema] = montosoloposdiaguardado(Date)
        [Forms]![Cierre por Turno]![deliverysistema] = montohugodiaguardado(Date)
        [Forms]![Cierre por Turno]![transferenciasistema] = montoposdiaguardadoTransferencia(Date)
        
        [Forms]![Cierre por Turno].Form.Requery
        
        'Nuevo Procedimiento
        DoCmd.OpenForm "CierreCorroborarDatosPOS"
        [Forms]![CierreCorroborarDatosPOS]![aCodOperario] = [Forms]![Main Pitaya]![codigologin]
        [Forms]![CierreCorroborarDatosPOS]![aHoraFinal] = Time()
        [Forms]![CierreCorroborarDatosPOS]![afecha] = Date
        [Forms]![CierreCorroborarDatosPOS]![aTotalPOSSistema] = [Forms]![Cierre por Turno]![possistema]
    Else
        MsgBox "No se cargaron los datos correctamente, prueba abrir cierre otra vez"
    End If
    
End If

Exit Sub
Nulo:
MsgBox "No se ha creado registro de cierre, vuelva a intentarlo"

End Sub

Private Sub Comando1715_Click()
DoCmd.OpenForm "Relacion de Productos Venta"
End Sub

Private Sub Comando1716_Click()
DoCmd.OpenReport "ValeDeDinero", acViewNormal, , "[Fecha]=#" & Date & "#"
End Sub

Private Sub Comando2137_Click()
DoCmd.OpenForm "Perdidas por Unidades"
End Sub

Private Sub Comando2141_Click()
'Modulo Despacho
Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)
'Call importarpreingresossistema0

DoCmd.OpenForm "HistorialPreIngresosLocal"
[Forms]![HistorialPreIngresosLocal].codigoope = Me.codigologin '
Call Forms("[HistorialPreIngresosLocal]").modosucursal
End Sub



Private Sub Comando1769_Click()


If cajainicial(Date) = 0 And Not CurrentProject.AllForms("Inicial").IsLoaded Then
    MsgBox "Ingresar caja inicial para poder facturar"
    DoCmd.OpenForm "Inicial"
    [Forms]![Inicial].Form.SetFocus
    Exit Sub
End If

Dim Orden As String
Orden = CrearNotaDePedido(3)

End Sub



Private Sub Comando1999_Click()

End Sub

Private Sub Comando2139_Click()
DoCmd.OpenForm "Ingreso Inventario Pitaya"
[Forms]![Ingreso Inventario Pitaya]![codigologin] = Me.codigologin



End Sub





Private Sub Comando2164_Click()
DoCmd.OpenForm "Menu PITAYA Global"
[Forms]![Menu PITAYA Global]![solovista] = 1
[Forms]![Menu PITAYA Global]![Comando538].Caption = "CERRAR"
End Sub

Private Sub Comando2192_Click()
DoCmd.OpenForm "Ingreso de Compras"
End Sub

Private Sub Comando2204_Click()
If Hour(Now()) > 18 Then
    MsgBox "La hora limite para solicitud de insumos locales ya cerro, contactarse con la lider"
Else
    DoCmd.OpenForm "SolicitudInsumosLocalesDia"
    [Forms]![SolicitudInsumosLocalesDia]![codoperariopedido] = Me.codigologin
End If
End Sub



Private Sub Comando2211_Click()
DoCmd.OpenReport "NuevoRegistroClub", acViewNormal
End Sub

Private Sub Comando2217_Click()

DoCmd.OpenForm "LogueoAutorizacion"
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(5, codigoLocal(), Date)

[Forms]![LogueoAutorizacion]![origenlogueo] = "ElegirTipoSalidaEfectivo"

End Sub







Private Sub Comando2243_Click()
'cierre de caja version 2
On Error GoTo Nulo

If MsgBox("Desea realizar cierre de turno?", vbYesNo, "CONFIRMACION DE CIERRE DE TURNO") = vbNo Then
    Exit Sub
End If


If (Time - UltimoCierre()) * 24 * 60 < 60 Then 'da resultado minutos desde ultimo cierre a ahora, 1 hora despues del ultimo cierre no es permitido
    If MsgBox("Existe un cierre realizado recientemente. Desea solicitar reinicio de cierre?", vbYesNo, "SOLICITAR REINICIO DE CIERRE") = vbYes Then
        DoCmd.OpenForm "SolicitudAnulacionCierre"
        [Forms]![SolicitudAnulacionCierre]![pedidoanular] = CierreFinal([Forms]![Main Pitaya]![fechasistema])
        [Forms]![SolicitudAnulacionCierre]![pedidooperario] = NombreOperario([Forms]![Main Pitaya]![codigologin])
    Else
        ' no solicita nada
    End If
    Exit Sub
End If

Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)

If existepreingresopendientefecha(Me.fechasistema) > 0 And Time() > #5:00:00 PM# Then
    MsgBox "Existen Preingresos que no han sido ingresados aun, corregir y dar ingreso a todos los ingresos pendientes de hoy para poder hacer cierre"
    Exit Sub
Else
    
    Call actualizarmontoguardadopedidosdia(Date)
    DoCmd.OpenReport "HistorialVentasxPedidoDia", acViewReport, , "[Pago]='POS' AND [Fecha]=#" & Date & "#"
    DoCmd.SelectObject acReport, "HistorialVentasxPedidoDia"
    DoCmd.PrintOut acSelection, 1, 1
    DoCmd.Close acReport, "HistorialVentasxPedidoDia"
    
    Dim pedidoscalculados As Double
    Dim pedidosguardados As Double
    
    Dim pedidosguardadospos As Double
    Dim pedidosguardadostrasnfer As Double
    Dim pedidosguardadospedidosya As Double
    Dim pedidosguardadosefectivo As Double
    
    pedidoscalculados = AcumuladoDia(Me.fechasistema)
    pedidosguardados = AcumuladoDiaGuardado(Me.fechasistema)
    pedidosguardadospos = montosoloposdiaguardado(Date)
    pedidosguardadostrasnfer = montoposdiaguardadoTransferencia(Date)
    pedidosguardadospedidosya = montohugodiaguardado(Date)
    pedidosguardadosefectivo = pedidosguardados - pedidosguardadospos - pedidosguardadostrasnfer - pedidosguardadospedidosya
    
    If pedidoscalculados = pedidosguardados Then ' Datos cargados correctamente
        
        
        'Nuevo preingreso
        Dim ultimoc As Date
        ultimoc = UltimoCierre()
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CierreDiario(HoraInicial, HoraFinal, Fecha, CodOperario)" & _
        " values (#" & Format(ultimoc, "hh:nn:ss") & "#, #" & Format(Time(), "hh:nn:ss") & "#, #" & Date & "#, " & [Forms]![Main Pitaya]![codigologin] & ")"
        DoCmd.SetWarnings True
        
        Dim cierrel As Long
        cierrel = CierreFinal(Date)


        DoCmd.OpenForm "Cierre por Turno 2"
        [Forms]![Cierre por Turno 2]![CodCierre] = cierrel
        [Forms]![Cierre por Turno 2]![aCodOperario] = [Forms]![Main Pitaya]![codigologin]
        [Forms]![Cierre por Turno 2]![aHoraInicial] = ultimoc
        [Forms]![Cierre por Turno 2]![aHoraFinal] = Time()
        [Forms]![Cierre por Turno 2]![afecha] = Date
        [Forms]![Cierre por Turno 2]![montoperiodo] = pedidoscalculados
    
        [Forms]![Cierre por Turno 2]![vposcalculado] = pedidosguardadospos
        [Forms]![Cierre por Turno 2]![vpedidosyacalculado] = pedidosguardadospedidosya
        [Forms]![Cierre por Turno 2]![vtrasnfercalculado] = pedidosguardadostrasnfer
        [Forms]![Cierre por Turno 2]![vefeccalculado] = pedidosguardadosefectivo
        
        [Forms]![Cierre por Turno 2].Form.Requery
        
        'DoCmd.OpenReport "Cierre por Turno 2", acViewNormal
        
        'Nuevo Procedimiento

        ' Abrir formulario "Contador Caja" en vista de formulario
        DoCmd.OpenForm "Contador Caja 2", acNormal
        [Forms]![Contador Caja 2]![titulo1] = "PRE-CIERRE DE CAJA " & Date
        [Forms]![Contador Caja 2]![titulo2] = "Cajero: " & NombreOperario([Forms]![Main Pitaya]![codigologin])
        [Forms]![Contador Caja 2]![titulo3] = "Tipo de Cambio " & tipocambio(Date)
        [Forms]![Contador Caja 2]![tipocambiov] = tipocambio(Date)
        [Forms]![Contador Caja 2]![codigocierrex] = cierrel
        [Forms]![Contador Caja 2]![ahoracierre] = Time()
        [Forms]![Contador Caja 2]![ahoracierreinicio] = ultimoc
        [Forms]![Contador Caja 2]![acodigocajero] = [Forms]![Main Pitaya]![codigologin]
        
        [Forms]![Main Pitaya].Form.Requery
    Else
        MsgBox "No se cargaron los datos correctamente, prueba abrir cierre otra vez"
    End If
    
End If

Exit Sub
Nulo:
MsgBox "No se ha creado registro nuevo de cierre, vuelva a intentarlo"
End Sub

Private Sub Comando2256_Click()

If cajainicial(Date) = 0 And Not CurrentProject.AllForms("Inicial").IsLoaded Then
    MsgBox "Ingresar caja inicial para poder facturar"
    DoCmd.OpenForm "Inicial"
    [Forms]![Inicial].Form.SetFocus
    Exit Sub
End If

Dim Orden As String
Orden = CrearNotaDePedido(4)
End Sub

Private Sub Comando2300_Click()
Call importartablaespecifica("Main_DB", "DBIngredientes", "DBIngredientes", 1)
DoCmd.OpenForm "LogueoAutorizacion"
'[Forms]![LogueoUsuario]![CodigoBusqueda] = Me.codigologin
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(21, 18, Date)
[Forms]![LogueoAutorizacion]![CodigoBusqueda].Enabled = True
[Forms]![LogueoAutorizacion]![origenlogueo] = "SeleccionInternaInsumosImportantes"

End Sub

Private Sub Comando2301_Click()
DoCmd.OpenForm "LogueoAutorizacion"
'[Forms]![LogueoUsuario]![CodigoBusqueda] = Me.codigologin
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(21, 18, Date)
[Forms]![LogueoAutorizacion]![CodigoBusqueda].Enabled = True
[Forms]![LogueoAutorizacion]![origenlogueo] = "AuditoriaCajaChica"
End Sub

Private Sub Comando2308_Click()
DoCmd.OpenForm "LogueoAutorizacion"
'[Forms]![LogueoUsuario]![CodigoBusqueda] = Me.codigologin
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(21, 18, Date)
[Forms]![LogueoAutorizacion]![CodigoBusqueda].Enabled = True
[Forms]![LogueoAutorizacion]![origenlogueo] = "IngresoAjustesInventario"
End Sub





Private Sub Comando2317_Click()
Call importartablaespecifica("Pitaya15_DB", "NotaDePedido", "NotaDePedidoDeliveryCentral", 3)
Call importartablaespecifica("Pitaya15_DB", "SubPedido", "SubPedidoDeliveryCentral", 3)
Call importartablaespecifica("Pitaya15_DB", "StatusPedidosCentral", "StatusPedidosCentralDeliveryCentral", 3)

DoCmd.OpenForm "HistorialVentasDeliveryCentral"
End Sub

Private Sub Comando2357_Click()
If MsgBox("Desea reinicar los datos del drive?", vbYesNo, "CONFIRMACION") = vbYes Then

    Call ReiniciarGoogleDrive
End If
End Sub

Private Sub Comando2366_Click()
Call actualizardatosclubmixedglobal
'Call DescargarTablaCompleta("clientesclub", "clientesclubexterno", "sucursal <> " & codigolocal())
End Sub

Private Sub Comando2376_Click()
If MsgBox("Desea subir los datos de exixstencias de un dia especifico?", vbYesNo, "CONFIRMACION") = vbYes Then
    Dim diaespecifico As Date
    diaespecifico = Date
    Call SyncKardexTiendaDiaEspecifico(diaespecifico)
    Call SyncVentasDia(diaespecifico)
    Call SyncLeerRespuestasAnulacion
    Call SyncAnularCierresPendientes
    MsgBox "Se actualizaron las existencias en la web correctamente"
End If
End Sub

Private Sub Comando671_Click()

DoCmd.OpenForm ("Clientes Club Pitaya")
End Sub

Private Sub Comando774_Click()
DoCmd.OpenForm "Ingreso Inventario Pitaya"
[Forms]![Ingreso Inventario Pitaya]![codigologin] = Me.codigologin
End Sub

Private Sub Comando91_Click()
DoCmd.OpenForm "IngresarCliente"
End Sub

Private Sub Form_Close()
If EsSistemaDeTienda() Then
    Call DetenerPingAutomatico
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error Resume Next
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.nombredelocal = UCase(nombrelocal())
Me.fechasistema = Date
Me.horaactual = Time()
Me.Requery

Me.usovidrio = PorcentajeUsoVidrio(Me.fechasistema)
Me.gigantonas = PorcentajeMedidasBatidos("Gigantona", Me.fechasistema)
Me.tipocambiosistema = "Cambio: " & tipocambio(Date)
'Me.Ventas = AcumuladoDia(Me.fechasistema)
Me.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Background\Background.jpg"

'1440 twips = 96 pixel = 1 inch
Me.InsideHeight = 11520
Me.InsideWidth = 14000
'MsgBox Me.WindowTop
'MsgBox Me.WindowLeft
Me.Move -645, -1260

If codigoLocal() = 15 Then ' sucursal central
    Me.Comando2256.Visible = True
    Me.Comando48.Visible = False 'pedido en tienda
    Me.Comando1769.Visible = False 'pedido en peiddosya
    Me.Imagen2281.Visible = False 'logo en peiddosya
    Me.Imagen2287.Visible = False 'logo en peiddosya
    Me.Comando2366.Visible = True 'boton de actualizar datos de sucursales
    Me.TipoPedido.width = 1600
    Me.nombresucursal.width = 1600
    Me.statuscentral.width = 1600
    Me.InsideWidth = 14000 + 3 * 1600
End If

  
Call AgregarTitulo  'Llamos al módulo para que nos ponga el título e icono al arrancar la aplicación.
Me.Imagen2281.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\PedidosYa.ico"
Me.Imagen2287.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\PedidosYa.ico"

If EsSistemaDeTienda() Then
    Call IniciarPingAutomatico
End If
End Sub

Private Sub Comando1217_Click()

DoCmd.OpenForm "Introduccion"
DoCmd.Close acForm, "Main Pitaya"
End Sub

Private Sub Comando1225_Click()
If Me.acodigo = "" Or IsNull(Me.acodigo) Then
    MsgBox "Ingresar un codigo valido"
    Exit Sub
End If

If APIDisponible() Then
    Call DescargarTablaCompleta("VentasGlobalesAccessCSV", "VentasGlobalesAccessCSVFiltradoCliente", "CodCliente = " & Me.acodigo & " AND local <> " & codigoLocal())
    Dim clientebuscar As Variant
    clientebuscar = DatosClienteClubGlobal(Me.acodigo)
    Call CargarVentasInternoExternoClienteHaciaTabla(Me.acodigo) 'crear tabla de istorial local y externo
    
    DoCmd.OpenForm "Historial Cliente 2"
    [Forms]![Historial Cliente 2]![acodigo] = Me.acodigo
    [Forms]![Historial Cliente 2]![anombre] = clientebuscar(1)
    [Forms]![Historial Cliente 2]![apuntos] = clientebuscar(0)
    [Forms]![Historial Cliente 2]![iniciales] = clientebuscar(3)
    [Forms]![Historial Cliente 2].Form.Requery
Else
    DoCmd.OpenForm "Historial Cliente"
    [Forms]![Historial Cliente]![acodigo] = Me.acodigo
    [Forms]![Historial Cliente]![anombre] = nombreCliente(Me.acodigo)
    [Forms]![Historial Cliente]![apuntos] = PuntosGlobales2(Me.acodigo)
    [Forms]![Historial Cliente].Form.Requery
End If
End Sub


Private Sub Comando1309_Click()
Me.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Background\Background" & Int((10 * Rnd) + 1) & ".jpg"
End Sub

Private Sub Comando1412_Click()
DoCmd.OpenForm "Control Porcionamiento"
[Forms]![Control Porcionamiento]![asemana] = numerosemana(Date)
End Sub

Private Sub Comando1587_Click()
DoCmd.OpenForm "Control Existencias Compra Venta"
End Sub





Private Sub Comando48_Click()

If cajainicial(Date) = 0 And Not CurrentProject.AllForms("Inicial").IsLoaded Then
    MsgBox "Ingresar caja inicial para poder facturar"
    DoCmd.OpenForm "Inicial"
    [Forms]![Inicial].Form.SetFocus
    Exit Sub
End If

Dim Orden As String
Orden = CrearNotaDePedido(1)
End Sub

Private Sub Comando50_Click()
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & Me.CodPedido

End Sub


Private Sub Form_Timer() 'Lee cada 30 segundos
 
On Error Resume Next
Dim Horal As Date
Horal = Time()
Me.horaactual = Horal

If EsSistemaDeTienda() Then  'que tenga la tabla DatosSistema vinculado y que no tenga nombre pitaya_system

    If cajainicial(Date) = 0 And Not CurrentProject.AllForms("Inicial").IsLoaded Then
        DoCmd.OpenForm "Inicial"
        [Forms]![Inicial].Form.SetFocus
    End If
    
    'Leer grupo de teegram mientras no este abierto nota de pedido
    If Not CurrentProject.AllForms("Nota de Pedido").IsLoaded And EsSistemaDeTienda() Then  ' si no esta abierto nota de pedido
        
        Call ComunicacionTelegram
        Call PingTimerTick
        Call SyncLeerRespuestasAnulacion
        Call SyncAnularCierresPendientes
        Me.Requery
    
    
        'If codigolocal() <> 15 Then
        '    Call importartablaespecifica("Pitaya15_DB", "StatusPedidosCentral", "StatusPedidosCentralDeliveryCentral", 3) 'descargar status de pedidos si hay
        '    Dim cantipedi As Integer
        '    cantipedi = cantidadpedidospendientescentral(codigolocal(), Date)
        '    If cantipedi > 0 Then
        '        Call NotificacionConSonido("pedidonuevo")
        '        Sleep 500
        '        Call NotificacionConSonido("pedidonuevo")
        '        Sleep 500
        '        MsgBox "HAY " & cantipedi & " PEDIDOS PENDIENDES POR FACTURAR DE LA CENTRAL", vbCritical
        '
        '        'Call importartablaespecifica("Pitaya15_DB", "NotaDePedido", "NotaDePedidoDeliveryCentral", 3)
        '        'Call importartablaespecifica("Pitaya15_DB", "SubPedido", "SubPedidoDeliveryCentral", 3)
        '        'Call importartablaespecifica("Pitaya15_DB", "StatusPedidosCentral", "StatusPedidosCentralDeliveryCentral", 3)
        '        'Call importartablaespecifica("Pitaya15_DB", "ClientesDelivery", "ClientesDelivery", 1)
        '        'DoCmd.OpenForm "HistorialVentasDeliveryCentral"
        '    End If
        '    DoCmd.RunSQL "DROP TABLE StatusPedidosCentralDeliveryCentral"
        'End If
    
    End If
    
    '================================================================================================
    'Fotos hora programada
    'If Hour(Horal) = 7 And Minute(Horal) = 15 And Second(Horal) < 30 Then 'Foto abierto local
    '    Call FotoTemporalTelegram
    'End If
    
    'If Hour(Horal) = 19 And Minute(Horal) = 30 And Second(Horal) < 30 Then 'Foto abierto local
    '    Call FotoTemporalTelegram
    'End If
    
    'If Hour(Horal) = 20 And Minute(Horal) = 30 And Second(Horal) < 30 Then 'Foto abierto local
    '    Call FotoTemporalTelegram
    'End If
    
    '================================================================================================
    ' CADA 60 MINUTOS
    If Minute(Horal) = 0 And Second(Horal) < 30 Then
        
        If Not CurrentProject.AllForms("Nota de Pedido").IsLoaded And EsSistemaDeTienda() Then    ' si no esta abierto nota de pedido
            Call ActualizarArchivoCSVVentasFecha(Date)
            Call ActualizarArchivoCSVVentasFecha(Date - 1)
        End If
        
    End If

End If
End Sub

Private Sub actualizarlista()
On Error GoTo NoResultados

Me.RecordSource = "SELECT NotaDePedido.CodPedido, NotaDePedido.Fecha, NotaDePedido.Hora, NotaDePedido.CodCliente," & _
" NotaDePedido.[Nombre Provicional], NotaDePedido.senuelo, MontoPedido([NotaDePedido]![CodPedido]) AS Monto," & _
" IIf([NotaDePedido]![Anulado]=0,'Activo','Anulado') AS Estado, NotaDePedido.Modalidad," & _
" IIf([NotaDePedido]![Modalidad]=0, 'PLASTICO','VIDRIO') AS Envase, IIf([NotaDePedido]![POS]=0,'EFECTIVO','TARJETA') AS Pago, " & _
" DLookUp('[Nombre]','[Delivery]','[CodDelivery]=' & [NotaDePedido]![Delivery]) AS Deliv" & _
" FROM NotaDePedido WHERE (((NotaDePedido.Fecha)=#" & Me.fechasistema & "#))" & _
" ORDER BY NotaDePedido.Fecha DESC , NotaDePedido.Hora DESC"

Exit Sub

NoResultados:
MsgBox "NO Data"


End Sub





