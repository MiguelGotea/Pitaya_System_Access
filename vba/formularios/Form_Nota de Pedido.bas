' ==========================================================
' Modulo  : Form_Nota de Pedido
' Tipo    : 100  |  Lineas: 987
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database
Public Sub BloquearFacturacion()
'MsgBox "Pedido Bloqueado,solo permitido cambiar tipo de envase y tipo de pago"
'Me.Secundario57.Locked = True 'ventana de productos facurados
Me.Secundario57.Form.CodPromocion.Enabled = False
Me.Secundario57.Form.Comando2050.Visible = False
Me.Secundario57.Form.Comando177.Visible = False
Me.Secundario57.Form.Comando202.Visible = False
Me.Secundario57.Form.eliminar.Visible = False
Me.Secundario57.Form.eliminar.Enabled = False
Me.Secundario57.Form.Observaciones.Locked = True
Me.Secundario57.Form.DetallesAdicionales.Locked = True

'Me.banulado.Enabled = False 'boton de anular
Me.CodCliente.Locked = True 'cambiar codigo de cliente bloqeuado
Me.[Nombre Provicional].Locked = True 'cambiar nombre opcional de cliente bloqeuado

Me.Comando1204.Enabled = False 'Boton agregar producto delivery
Me.Comando68.Enabled = False 'Boton agregar producto normal

Me.Comando155.Enabled = False 'Boton imprimir factura total
Me.tipopropina.Enabled = False

Me.CodMotorizado.Enabled = False

Me.rotulobloquear.Visible = True

'Boton de imprimir ya esta bloqueada pero se hace una consulta adicional para abilitara
If (Me.Impresiones < 2 Or IsNull(Me.Impresiones)) And Date = Me.Fecha And Me.Anulado = 0 Then ' si hay menos de 2 impresiones puede imprimir aun
    Comando155.Enabled = True
End If
End Sub
Public Sub DesbloquearFacturacion()
'MsgBox "Pedido Bloqueado,solo permitido cambiar tipo de envase y tipo de pago"
'Me.Secundario57.Locked = False 'ventana de productos facurados
Me.Secundario57.Form.CodPromocion.Enabled = True
Me.Secundario57.Form.Comando2050.Visible = True
Me.Secundario57.Form.Comando177.Visible = True
Me.Secundario57.Form.Comando202.Visible = True
Me.Secundario57.Form.eliminar.Visible = True
Me.Secundario57.Form.eliminar.Enabled = True
Me.Secundario57.Form.Observaciones.Locked = False
Me.Secundario57.Form.DetallesAdicionales.Locked = False
Me.tipopropina.Enabled = True

'Me.banulado.Enabled = True 'boton de anular
Me.CodCliente.Locked = False 'cambiar codigo de cliente bloqeuado
Me.[Nombre Provicional].Locked = False 'cambiar nombre opcional de cliente bloqeuado

Me.Comando1204.Enabled = True 'Boton agregar producto delivery
Me.Comando68.Enabled = True 'Boton agregar producto normal

Me.Comando155.Enabled = True 'Boton imprimir factura total
Me.CodMotorizado.Enabled = True ' eleccion de motorizado
Me.bpos.Enabled = True ' boton de cambiar tipo de pago

If IsNull(Me.CodMotorizado) Then ' elecciond e mtorizado
    Me.CodMotorizado.Enabled = False
Else
    Me.CodMotorizado.Enabled = True
End If

Me.rotulobloquear.Visible = False

Me.Secundario57.SetFocus

End Sub

Private Sub anularpedido_Click()
Dim sustento As String

If Date <> Me.Fecha Then ' CUanado se trata de eliminar un pedido de otra fecha, no realiza ninguna accion
    'No hace nada
Else
    If (Time > Me.Hora + 10 / 60 / 24 Or Me.Impreso <> 0) And codigoLocal() <> 15 Then ' si se hace la solciitiud fuera de tiempo permitido
        
        If Me.banulado.Caption = "ACTIVO" Then
        
            ' cuando el pedido ya esta fuera de tiempo y solo cuando no esta anulado
            If existesolicitudanulacionpedido(Me.CodPedido) = 0 Then ' no existe solicitud de anualcion en curso
                DoCmd.OpenForm "SolicitudAnulacionPedido"
                
                [Forms]![SolicitudAnulacionPedido]![Texto1305].Caption = "SOLICITUD DE ANULACION DE PEDIDO"
                [Forms]![SolicitudAnulacionPedido]![Etiqueta1289].Visible = True
                [Forms]![SolicitudAnulacionPedido]![pedidocambio].Visible = True
                
                [Forms]![SolicitudAnulacionPedido]![pedidoanular] = Me.CodPedido
                [Forms]![SolicitudAnulacionPedido]![pedidooperario] = NombreOperario(Me.colaborador)
            Else
                MsgBox "Ya se solicito anulacion de este pedido, revisar status de anulacines"
            End If
        Else
            
            ' SI el pedido esta anulado ya no s epuede pedir desanular despues de tiempo
        
        End If
        
    Else  ' dentro de tiempo permitido y cuando no ha sido impreso y si es local 15
        
        ' cuado el pedido no esta bloqeuado se anula noral sin ningun problema
        If Me.Anulado.Value = False Then   'Cuando esta activo pasa a anulado
    
            DoCmd.OpenForm "SolicitudAnulacionPedido"
            
            [Forms]![SolicitudAnulacionPedido]![Texto1305].Caption = "ANULACION DE PEDIDO"
            [Forms]![SolicitudAnulacionPedido]![Etiqueta1289].Visible = False
            [Forms]![SolicitudAnulacionPedido]![pedidocambio].Visible = False
            
            [Forms]![SolicitudAnulacionPedido]![pedidoanular] = Me.CodPedido
            [Forms]![SolicitudAnulacionPedido]![pedidooperario] = NombreOperario(Me.colaborador)
    
        Else                               'Cuando pasa de anulado a activo, y borra el registro de anulado que se hiso
            
            If codigoLocal() = 15 Then 'ya no puede desanular
                'esta habilitado el boton de desanular pero si e sla central ya no se puede
                MsgBox "Pedido anulado, ya no se puede desbloquear, facturar nuevamente"
            Else
                Me.Anulado.Value = False
                Me.MotivoAnulado.Visible = False
                Me.EtiquetaMotivoAnulado.Visible = False
                Me.banulado.Caption = "ACTIVO"
                Me.banulado.BackColor = RGB(34, 177, 76)
                
                DoCmd.SetWarnings False
                DoCmd.RunSQL "DELETE * FROM AnulacionPedidos WHERE CodPedido = " & Me.CodPedido
                DoCmd.SetWarnings True
                
                Call DesbloquearFacturacion
            End If
            
        End If
        
    End If
End If
End Sub

Private Sub botonaprobar_Click()
If Me.Anulado.Value = True Then
    'Pedido anulado no se puede aprobar
    MsgBox "No se puede enviar un pedido anulado a la sucursal"
    Exit Sub
End If

If Me.CodClientesDelivery = 0 Or IsNull(Me.CodClientesDelivery) Then
    MsgBox "Ingresar datos de envio"
    Exit Sub
End If

If statusaprobaciondeliverycentral(Me.CodPedido) = True Then
    MsgBox "El pedido ya se encuentra enviado a la sucursal"
    Exit Sub
End If

Dim desti As Integer
Dim pedix As Long
Dim nombrebot As String

desti = DLookup("[CodLocal]", "[ClientesDelivery]", "[CodClientesDelivery]=" & Me.CodClientesDelivery)
nombrebot = DLookup("[BotTelegram]", "[StatusSucursales]", "[CodLocal]=" & desti)
pedix = Me.CodPedido

If MsgBox("Desea enviar el pedido a la sucursal?", vbYesNo, "CONFIRMACION") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO statusPedidosCentral(CodPedidoCentral, HoraAprobadoCentral, AprobadoCentral, Sucursal, Fecha)" & _
    " values (" & Me.CodPedido & ", #" & Now() & "#, -1," & desti & ", Date)"
    DoCmd.SetWarnings True
    Call EnviarMensajeTelegram(nombrebot & Chr(32) & "PedidoAprobado " & pedix & vbCrLf & _
    "Pedido: " & pedix & vbCrLf & _
    "Sucursal: " & nombrelocalglobal(desti) & vbCrLf & _
    "Servicio: " & DLookup("[TipoPedido]", "[ClientesDelivery]", "[CodClientesDelivery]=" & Me.CodClientesDelivery) & vbCrLf & _
    detalletextopedido(pedix), grupotpedidospitayacentral())
    MsgBox "Pedido enviado correctamente a la sucursal"
End If
End Sub

Private Sub bpos_Click()
'TIPO DE PAGO CAMBIABLE APLICA SOLO PEDIDOS CON MOTORIZADO PROPIO O DE CLIENTES, PEDIODYA Y HUGO SE BLOQUEA OPCION
'rotacion  EFECTIVO -> TRANSFERENCIA - > POS -> PEDIDOSYA -> HUGO
Dim captionOriginal As String
captionOriginal = Me.bpos.Caption

Select Case captionOriginal

Case "EFECTIVO"
    Me.POS = -1
    Me.Transferencia = -1
    Me.DeliveryComision = 0
    Me.bpos.Caption = "TRANSFERENCIA" 'siguiente boton
    'Me.TipoDelivery.Enabled = False
    
Case "TRANSFERENCIA"
    Me.POS = -1
    Me.Transferencia = 0
    Me.DeliveryComision = 0
    Me.bpos.Caption = "POS" 'siguiente boton
    'Me.TipoDelivery.Enabled = False

    
'Case "DELIVERY"
'    Me.POS = 0
'    Me.Transferencia = 0
'    Me.DeliveryComision = 0
'    Me.bpos.Caption = "EFECTIVO" 'siguiente boton
'    'Me.TipoDelivery.Enabled = False

Case "POS"
'    Me.POS = -1
'    Me.Transferencia = 0
'    Me.DeliveryComision = -1
'    Me.bpos.Caption = "DELIVERY" 'siguiente boton
'    'Me.TipoDelivery.Enabled = True
    
    Me.POS = 0
    Me.Transferencia = 0
    Me.DeliveryComision = 0
    Me.bpos.Caption = "EFECTIVO" 'siguiente boton
    'Me.TipoDelivery.Enabled = False
    
End Select

End Sub

Public Sub cargardatosclubfactura()
Dim nombrecli As String
Dim proce As Integer


If IsNull(Me.CodCliente) Then
    MsgBox "Ingresar codigo decliente valido o 0"
    Me.CodCliente = 0
    
    Me.Encabezado_automático0 = "NOTA DE PEDIDO " & Me.CodPedido
    Me.Cuadro_combinado123 = ""
    Me.ppedido.Requery
    Me.pacumulado = 0
    Me.clubguardado = 0
    Me.piniciales = 0
End If

If Me.CodCliente = 0 Then
    Me.Nombre_Provicional.Enabled = True
    
    Me.Encabezado_automático0 = "NOTA DE PEDIDO " & Me.CodPedido
    Me.Cuadro_combinado123 = ""
    Me.ppedido.Requery
    Me.pacumulado = 0
    Me.clubguardado = 0
    Me.piniciales = 0
    
Else
    If Me.CodCliente = Me.clubguardado Then
        'ya habia guardado eso datos asi que no carga nada
    Else
        If APIDisponible() Then 'cuando hay conexion a la api
            Dim clientebuscar As Variant
            clientebuscar = DatosClienteClubGlobal(Me.CodCliente)
                
            If clientebuscar(2) = 1 Then ' si existe el club local o en host
                Me.Nombre_Provicional = clientebuscar(1)
                Me.Nombre_Provicional.Enabled = False
                
                Me.Encabezado_automático0 = "NOTA DE PEDIDO " & Me.CodPedido & " - " & UCase(clientebuscar(1))
                Me.Cuadro_combinado123 = clientebuscar(1)
                'Me.piniciales = PuntosGlobales2(Me.CodCliente)
                Me.piniciales = clientebuscar(0)
                Me.pacumulado = IIf(Me.piniciales < 0, 0, Me.piniciales) '- Me.ppedido  'el sistema calcula clientebiuscar pero los productos ya aun no se le calcula susu putnos por producto asi que no sumara nada en el total
                Me.clubguardado = Me.CodCliente
                Me.ppedido.Requery
                    
             Else
                Me.CodCliente = 0
                
                Me.Encabezado_automático0 = "NOTA DE PEDIDO " & Me.CodPedido
                Me.Cuadro_combinado123 = ""
                Me.ppedido.Requery
                Me.pacumulado = 0
                Me.clubguardado = 0
                Me.piniciales = 0
                
                MsgBox "Codigo de cliente no se encuentra registrado, verificar numero de membresia y reportar al supervisor"
                
             End If
        Else 'cuando no hay conexion a la api
            If confirmarclub(Me.CodCliente) = 0 Then
                Me.CodCliente = 0
                
                Me.Encabezado_automático0 = "NOTA DE PEDIDO " & Me.CodPedido
                Me.Cuadro_combinado123 = ""
                Me.ppedido.Requery
                Me.pacumulado = 0
                Me.clubguardado = 0
                Me.piniciales = 0
                
                MsgBox "No hay conexion a internet,verificar wifi de computadora y reportar al area de sistemas"
            Else ' existe club
                Me.Nombre_Provicional = ""
                Me.Nombre_Provicional.Enabled = False
                nombrecli = UCase(nombreCliente(Me.CodCliente))
                Me.Nombre_Provicional = nombrecli
                
                Me.Encabezado_automático0 = "NOTA DE PEDIDO " & Me.CodPedido & " - " & nombrecli
                Me.Cuadro_combinado123 = nombrecli
                Me.Requery
                Me.ppedido.Requery
                Me.piniciales = PuntosGlobales2(Me.CodCliente)
                Me.pacumulado = Me.piniciales - Me.ppedido
                Me.clubguardado = Me.CodCliente
            End If
        End If
    End If
End If

Me.Requery
[Forms]![Nota de Pedido].Form.ppedido.Requery
[Forms]![Nota de Pedido].Form.Texto193.Requery
Me.Secundario57.Form.CodPromocion.Requery
Me.Secundario57.Requery


End Sub



Private Sub CodCliente_Exit(Cancel As Integer)
Call cargardatosclubfactura
End Sub

Private Sub CodMotorizado_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Comando1204_Click()
Me.Texto166 = 0
DoCmd.OpenForm "Menu PITAYA Delivery"
End Sub







Private Sub Comando1396_Click()
DoCmd.OpenForm "IngresarCliente"
End Sub

Private Sub Comando155_Click() 'Imrpimir boleta
Dim aleato As Integer
Dim codipedido As Long
Dim impresionespedido As Integer
Dim puntosonlycanjeados As Double
codipedido = Me.CodPedido
impresionespedido = 0
[Forms]![Nota de Pedido]![horaabierto] = Now

'''''''''''''ACTUALIZAR Y VALIDAR PUNTOS CLUB ANTES DE IMPRIMIR O MODIFICAR DATOS'''''''''''''
Call ActualizarPuntosSubPedidoDePedido(codipedido)

puntosonlycanjeados = -1 * puntosclubsolocanjeofactura(codipedido) 'la funcion devuelve negativo, se convierte apra validar todo en canjeado con disponibles
If puntosonlycanjeados <> 0 Then
    If puntosonlycanjeados > Nz(Me.pacumulado, 0) Then
        MsgBox "El cliente no cuenta con suficientes puntos acumulados para concretar este pedido." & vbCrLf & vbCrLf & _
               "Puntos requeridos para canje: " & puntosonlycanjeados & vbCrLf & _
               "Puntos acumulados disponibles: " & Nz(Me.pacumulado, 0), vbExclamation, "Puntos Insuficientes"
        Exit Sub
    End If
End If

'''''''''''''VERIFICACION DE NOMBRE DE LCIENTE ''''''''
If Me.CodCliente = 0 And (IsNull(Me.[Nombre Provicional]) = True Or Me.[Nombre Provicional] = "" Or Me.[Nombre Provicional] = " ") Then
    Me.[Nombre Provicional] = IngresarNombre()
End If

If codigoLocal() <> 15 Then ' aplica si no es central cuando funcionaba pedido desde la central

'''''VERIFICACION DE MOTORIZADO
'    If pedidocontieneproducto("Servicio Delivery", codipedido) <> 0 And IsNull(Me.CodMotorizado) Then ' hay costo de envio facturado
'        MsgBox "Ingresar nombre de motorizado para imprimir factura"
'        Exit Sub
'    End If
'    If (IsNull(Me.CodMotorizado) = False Or Me.CodMotorizado <> "") And pedidocontieneproducto("Servicio Delivery", codipedido) = 0 Then ' hay mtorizado pero no se facturo producto
'        MsgBox "Ingresar Costo de envio para imprimir factura"
'        Exit Sub
'    End If

'''''VERIFICAION DE IMPRESIONES DE FACTURA ''''''''''
    If Me.Impreso <> -1 Then ' aun no  se imprimio
        Me.HoraImpreso = Time
        Me.Hora = Time
    End If
    Me.Impreso = -1
    
    If IsNull(Me.Impresiones) Then
        Me.Impresiones = 0
    End If
    
    Me.Impresiones = Me.Impresiones + 1
    impresionespedido = Me.Impresiones

End If

''''''''''''''''REFRESCAR DATOS '''''''''
Me.Requery
Sleep 1000
[Forms]![Nota de Pedido]![horaabierto] = Now

'''''''''''''QUE IMPRIME CENTRAL ''''''''''
'If codigoLocal() = 15 Then   ' solo aplica sucursal central
'    If Me.CodClientesDelivery = 0 Or IsNull(Me.CodClientesDelivery) Then
'        MsgBox "Ingresar datos de envio"
'        Exit Sub
'    End If
'
'    DoCmd.OpenForm "BoletaExclusivoCentral", acNormal, , "[CodPedido]=" & codipedido
'
'    DoCmd.MoveSize , 0, , GetSystemMetrics(1) * 15 - 3750 ' altura de pantalla pixeles a twins - atura de taskbar 250pixeles
'
'    [Forms]![BoletaExclusivoCentral]![ACodLocal] = DLookup("[Nombre]", "[StatusSucursales]", "[CodLocal]=" & DLookup("[CodLocal]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery]))
'    [Forms]![BoletaExclusivoCentral]![aTelefono] = DLookup("[Telefono]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery])
'    [Forms]![BoletaExclusivoCentral]![aDireccion] = DLookup("[Direccion]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery])
'    [Forms]![BoletaExclusivoCentral]![aTipoPedido] = DLookup("[TipoPedido]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery])
'    If [Forms]![BoletaExclusivoCentral]![aTipoPedido] = "Delivery" Then
'        [Forms]![BoletaExclusivoCentral]![aCodProovedoresDelivery] = DLookup("[Nombre]", "[ProovedoresDelivery]", "[CodProovedoresDelivery]=" & DLookup("[CodProovedoresDelivery]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery]))
'        If [Forms]![BoletaExclusivoCentral]![aCodProovedoresDelivery] = "Motorizado Pitaya" Then
'            [Forms]![BoletaExclusivoCentral]![aDistancia] = DLookup("[Distancia]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery])
'            [Forms]![BoletaExclusivoCentral]![aConductor] = NombreOperario(DLookup("[CodigoMotorizado]", "[ClientesDelivery]", "[CodClientesDelivery]=" & [Forms]![Nota de Pedido]![CodClientesDelivery]))
'        End If
'    End If
'    Exit Sub
'End If

'''''QUE IMPRIME SUCURSALES  SOLO APLICA SUCURSALES POR DEFECTO TODO AHCIA ADELANTE ''''''''
Dim ciudadlocal As String
ciudadlocal = DLookup("[Ciudad]", "[StatusSucursales]", "[CodLocal]=" & codigoLocal())

If Me.TipoDelivery = 8 And ciudadlocal = "Managua" Then ' cuando es pedidosya imprime la otra boleta exclusiva
    DoCmd.OpenForm "BoletaExclusivoPedidosYa", acNormal, , "[CodPedido]=" & codipedido
    DoCmd.PrintOut
    DoCmd.Close acForm, "BoletaExclusivoPedidosYa"

Else
    'Dim puntos_sorteo As Integer
    
    'Dim tipo_cliente_sorteo As Integer
    'Dim puntos_ventas As Integer
    'Dim puntos_semilla As Integer
    'tipo_cliente_sorteo = IIf(Me.CodCliente <> 0, 2, 1)
    'puntos_ventas = Int(Me.TotalTotal / 250)
    'puntos_semilla = cantidadsubgrupodegrupopedido(Me.CodPedido, 7, 1)
    'puntos_sorteo = (puntos_ventas + puntos_semilla) * tipo_cliente_sorteo
    DoCmd.OpenForm "Boleta", acNormal, , "[CodPedido]=" & codipedido
    
    'If puntos_sorteo > 0 Then
    '    Call GenerarQRFactura(Me.CodPedido, Me.TotalTotal, puntos_sorteo, codigoLocal())
    '    [Forms]![Boleta]![Etiqueta405].Visible = True
    '    [Forms]![Boleta]![codigo_sorteo].Visible = True
    '    [Forms]![Boleta]![codigo_sorteo] = Me.CodPedido
    '    [Forms]![Boleta]![Etiqueta415].Visible = True
    '    [Forms]![Boleta]![puntos_sorteo].Visible = True
    '    [Forms]![Boleta]![puntos_sorteo] = puntos_sorteo
    '    [Forms]![Boleta]![qr_sorteo].Visible = True
    '    [Forms]![Boleta]![qr_sorteo].Picture = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\" & "QR_" & Me.CodPedido & ".png"
    'End If
        
    '[Forms]![Boleta].Form.Requery
    
    DoCmd.PrintOut
    DoCmd.Close acForm, "Boleta"
    
    'If puntos_sorteo > 0 Then
    '    Call LimpiarQRFactura(Me.CodPedido)
    'End If
    
    'If pedidocontieneproducto("Servicio Delivery", codipedido) <> 0 Then  ' hat delivery propio imprime copia a mtorizado, se imprime boleta adicional
    '    DoCmd.OpenForm "Boleta", acNormal, , "[CodPedido]=" & codipedido
    '    '[Forms]![Boleta]![rotulopedido].Visible = True
    '    [Forms]![Boleta]![CodPedido].Visible = True
    '    [Forms]![Boleta]![wificliente].Visible = False
    '    [Forms]![Boleta]![despedida].Visible = False
    '    [Forms]![Boleta]![copiamotorizado].Visible = True
    '
    '    [Forms]![Boleta]![titulonombremotorizado].Visible = True
    '    [Forms]![Boleta]![nombremotorizado].Visible = True
    '    [Forms]![Boleta]![nombremotorizado] = NombreOperario([Forms]![Nota de Pedido]![CodMotorizado])
    '    DoCmd.PrintOut
    '    DoCmd.Close acForm, "Boleta"
    'End If
End If

''''''''''''''''''''''''''''''''''''''''CUPONESS
'If Me.TipoDelivery = 0 And pedidocontienecuponusado(Me.CodPedido) = 0 Then
'
'    If pedidocontienebowlowaffle(Me.CodPedido) = -1 Then
'        [Reports]![Boleta]![cupon].Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Cupones\1.jpg"
'    ElseIf pedidocontienebatidoolimonada(Me.CodPedido) = -1 Then
'        aleato = Int(2 + Rnd * (5))
'        'Int(Límite inferior + NúmAleat*(Límite superior- Límite inferior + 1))
'        [Reports]![Boleta]![cupon].Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Cupones\" & aleato & ".jpg"
'    Else
'        [Reports]![Boleta]![cupon].Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Cupones\0.jpg"
'    End If
'End If


'''''''''''''''''CONTROL DE IMPRESIONES'''''''''''''
[Forms]![Nota de Pedido]![horaabierto] = Now
If impresionespedido = 1 Then
    
    'If codigolocal() <> 12 Or codigolocal() <> 7 Or codigolocal() <> 17 Then ' comanda detallada de despacho para otras sucursales
    '    DoCmd.OpenForm "ComandaEntregaPedido", acNormal, , "[CodPedido]=" & codipedido
    '    DoCmd.PrintOut
    '    DoCmd.Close acForm, "ComandaEntregaPedido"
    'End If
    
    'If ConcursoActivo() = True Then DoCmd.OpenReport "TicketConcurso", acViewNormal
    
    'If codigolocal() = 12 Or codigolocal() = 7 Or codigolocal() = 17 Then
    
        'altura de observaciones texto encabezaado 2.15 3096 actual
        ' altura de contenido nombrecondi 1.6 2304 actual  230/linea
        ' altura de endulzante defecto 0.3" 432
        '2 lineas exxtras por espacios entre grupo y grupo
        Dim cantilineas As Integer
        
        cantilineas = cantidadingredientesmaximorecetavisiblespedido(Me.CodPedido, 1)
        If cantilineas <> 0 Then
            DoCmd.OpenForm "ComandaEstacion", acDesign
            Forms("ComandaEstacion").Texto393.height = 0 ' 0 para no mostrar en segunda version de comanda 432 ' default 0.3" 432
            Forms("ComandaEstacion").nombrecondi.height = 50 + cantilineas * 225 + 2 * 225 + 2 * 255 ' 2 lineas mas de endulzante que pasa a este grupo
            Forms("ComandaEstacion").Section(acDetail).height = 3096 + 100 + cantilineas * 225 + 2 * 225 ' 4680 '5400actual  ' Altura en twips (1 cm ˜ 567 twips)
            DoCmd.Close acForm, "ComandaEstacion", acSaveYes
            
            DoCmd.OpenForm "ComandaEstacion", acNormal, , "[CodPedido]=" & codipedido & " AND [EstacionTrabajo]=1"
            [Forms]![ComandaEstacion]![titulo].Caption = "BATIDOS Y LIMONADAS"
            [Forms]![ComandaEstacion]![aestacion] = 1
            DoCmd.PrintOut
            DoCmd.Close acForm, "ComandaEstacion"
            'Batidos maximo 7 visible fit ptoeinico
        End If
        
        cantilineas = cantidadingredientesmaximorecetavisiblespedido(Me.CodPedido, 2)
        If cantilineas <> 0 Then
            DoCmd.OpenForm "ComandaEstacion", acDesign
            Forms("ComandaEstacion").Texto393.height = 0 ' default 0.3" 432
            Forms("ComandaEstacion").nombrecondi.height = 50 + cantilineas * 225 + 2 * 225 + 50
            Forms("ComandaEstacion").Section(acDetail).height = 3096 + 100 - 432 + cantilineas * 225 + 2 * 225 ' 4680 '5400actual  ' Altura en twips (1 cm ˜ 567 twips)
            DoCmd.Close acForm, "ComandaEstacion", acSaveYes
            
            DoCmd.OpenForm "ComandaEstacion", acNormal, , "[CodPedido]=" & codipedido & " AND [EstacionTrabajo]=2"
            [Forms]![ComandaEstacion]![titulo].Caption = "BOWLS, PARFAIT Y WAFFLES"
            [Forms]![ComandaEstacion]![aestacion] = 2
            DoCmd.PrintOut
            DoCmd.Close acForm, "ComandaEstacion"
            'Bowl maximo 10 visible acai bowl 9 dragon, waffle 7
        End If
    'Else
    '    Call LoopCocina(codipedido)
    'End If

End If

' '''''''''''''''CERRAR'''''''''''4
'Acumulado Dia
'Dim rut1 As String: rut1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Diario\" & nombrelocal() & "\Ventas Acumuladas.txt"
'Open rut1 For Output As #1
'Print #1, nombrelocal() & " = " & AcumuladoDia(Date)
'Close #1

'If Me.TipoDelivery <> 0 Then
'    'Ventas Hugo
'    Dim rut2 As String: rut2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Diario\" & nombrelocal() & "\Delivery.txt"
'    Open rut2 For Output As #1
'    Print #1, "Piki:  " & CantidadxTipoDelivery(Date, 6) & vbCrLf & _
'              "Hugo:  " & CantidadxTipoDelivery(Date, 5) & vbCrLf & _
'              "Pedidosya:  " & CantidadxTipoDelivery(Date, 8) & vbCrLf & _
'              "Ruta1: " & CantidadxTipoDelivery(Date, 1) & vbCrLf & _
'              "Ruta2: " & CantidadxTipoDelivery(Date, 2) & vbCrLf & _
'              "Ruta3: " & CantidadxTipoDelivery(Date, 3) & vbCrLf & _
'              "Ruta4: " & CantidadxTipoDelivery(Date, 4)
'    Close #1
'End If

If CurrentProject.AllForms("Main Pitaya").IsLoaded Then
    [Forms]![Main Pitaya].Form.Requery
End If
If CurrentProject.AllForms("HistorialVentasFiltro").IsLoaded Then
    [Forms]![HistorialVentasFiltro].Form.Requery
End If
If CurrentProject.AllForms("HistorialVentasPedido").IsLoaded Then
    [Forms]![HistorialVentasPedido].Form.Requery
End If


'''''''GUARDAR MONTO TOTAL ''''''
Me.TotalGuardado = MontoPedido(codipedido)
Me.Requery
'Call ActualizarPuntosSubPedidoDePedido(codipedido) 'Se ejecuta al inicio de Comando155_Click antes de imprimir


If EsSistemaDeTienda() = True Then
    
    If [Forms]![Nota de Pedido]![CodCliente] <> 0 Then ' club enviar notificacion si hay uso de puntos
        If puntosonlycanjeados <> 0 Then
            Call enviarnotificacionclienteusopuntos([Forms]![Nota de Pedido]![CodCliente], puntosonlycanjeados, nombrelocalglobal(codigoLocal()))
            MsgBox "Se ha notificado al cliente sobre el uso de sus puntos"
        End If
    End If

    Call SyncVentasPedido([Forms]![Nota de Pedido]![CodPedido]) ' enviar dato individual de pedido
End If

DoCmd.Close acForm, "Nota de Pedido"

End Sub

Private Sub Comando852_Click()
Me.Requery
End Sub


Private Sub Comando1658_Click()

' '''''''''''''''CERRAR'''''''''''4


'Acumulado Dia
'Dim rut1 As String: rut1 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Diario\" & nombrelocal() & "\Ventas Acumuladas.txt"
'Open rut1 For Output As #1
'Print #1, nombrelocal() & " = " & AcumuladoDia(Date)
'Close #1

'If Me.TipoDelivery <> 0 Then
'    'Ventas Hugo
'    Dim rut2 As String: rut2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Diario\" & nombrelocal() & "\Delivery.txt"
'    Open rut2 For Output As #1
'    Print #1, "Piki:  " & CantidadxTipoDelivery(Date, 6) & vbCrLf & _
'              "Hugo:  " & CantidadxTipoDelivery(Date, 5) & vbCrLf & _
'              "Pedidosya:  " & CantidadxTipoDelivery(Date, 8) & vbCrLf & _
'              "Ruta1: " & CantidadxTipoDelivery(Date, 1) & vbCrLf & _
'              "Ruta2: " & CantidadxTipoDelivery(Date, 2) & vbCrLf & _
'              "Ruta3: " & CantidadxTipoDelivery(Date, 3) & vbCrLf & _
'              "Ruta4: " & CantidadxTipoDelivery(Date, 4)
'    Close #1
'End If
Me.Impreso = True
Me.Requery

Me.TotalGuardado = MontoPedido(Me.CodPedido)
Me.Requery
Call ActualizarPuntosSubPedidoDePedido(Me.CodPedido)

If CurrentProject.AllForms("Main Pitaya").IsLoaded Then
    [Forms]![Main Pitaya].Form.Requery
End If
If CurrentProject.AllForms("HistorialVentasFiltro").IsLoaded Then
    [Forms]![HistorialVentasFiltro].Form.Requery
End If
If CurrentProject.AllForms("HistorialVentasPedido").IsLoaded Then
    [Forms]![HistorialVentasPedido].Form.Requery
End If
If CurrentProject.AllForms("BoletaExclusivoCentral").IsLoaded Then
    DoCmd.Close acForm, "BoletaExclusivoCentral"
End If

If EsSistemaDeTienda() Then
    Call SyncVentasPedido([Forms]![Nota de Pedido]![CodPedido]) ' enviar dato individual de pedido
End If
DoCmd.Close acForm, "Nota de Pedido"
End Sub

Private Sub Comando1661_Click()
DoCmd.OpenForm "Clientes Club Pitaya"
End Sub

Private Sub Comando1669_Click()
If Me.CodClientesDelivery = 0 Or IsNull(Me.CodClientesDelivery) Then
    DoCmd.OpenForm "IngresarClienteDelivery"
    [Forms]![IngresarClienteDelivery]![tiporegistro] = 0
Else
    DoCmd.OpenForm "IngresarClienteDelivery"
    Dim curso As Long
    curso = [Forms]![Nota de Pedido]![CodClientesDelivery]
    [Forms]![IngresarClienteDelivery]![tiporegistro] = 1
    [Forms]![IngresarClienteDelivery]![codigoencurso] = curso
    
    [Forms]![IngresarClienteDelivery]![ACodLocal] = DLookup("[CodLocal]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aTelefono] = DLookup("[Telefono]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aDireccion] = DLookup("[Direccion]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aTipoPedido] = DLookup("[TipoPedido]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aCodProovedoresDelivery] = DLookup("[CodProovedoresDelivery]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aCodigoMotorizado] = DLookup("[CodigoMotorizado]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aCostoDelivery] = DLookup("[CostoDelivery]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)
    [Forms]![IngresarClienteDelivery]![aDistancia] = DLookup("[Distancia]", "[ClientesDelivery]", "[CodClientesDelivery]=" & curso)

End If
End Sub

Private Sub Comando1752_Click()
DoCmd.OpenForm "Menu Cupones Web"
End Sub

Private Sub Comando673_Click()
If APIDisponible() Then

    If Me.CodCliente <> 0 Then
        Call DescargarTablaCompleta("VentasGlobalesAccessCSV", "VentasGlobalesAccessCSVFiltradoCliente", "CodCliente = " & Me.CodCliente & " AND local <> " & codigoLocal())
        Dim clientebuscar As Variant
        clientebuscar = DatosClienteClubGlobal(Me.CodCliente)
        Call CargarVentasInternoExternoClienteHaciaTabla(Me.CodCliente) 'crear tabla de istorial local y externo
        
        DoCmd.OpenForm "Historial Cliente 2"
        [Forms]![Historial Cliente 2]![acodigo] = Me.CodCliente
        [Forms]![Historial Cliente 2]![anombre] = clientebuscar(1)
        [Forms]![Historial Cliente 2]![apuntos] = clientebuscar(0)
        [Forms]![Historial Cliente 2]![iniciales] = clientebuscar(3)
        [Forms]![Historial Cliente 2].Form.Requery
        
    End If
Else
    If Me.CodCliente <> 0 Then
        DoCmd.OpenForm "Historial Cliente"
        [Forms]![Historial Cliente]![acodigo] = Me.CodCliente
        [Forms]![Historial Cliente]![anombre] = nombreCliente(Me.CodCliente)
        [Forms]![Historial Cliente]![apuntos] = Me.pacumulado
        [Forms]![Historial Cliente].Form.Requery
    End If
End If
End Sub

Private Sub Comando68_Click()
Me.Texto166 = 0
DoCmd.OpenForm "Menu PITAYA Global"
End Sub




Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

Dim clie As String
Dim puntosglobales As Double
         
''DATOS DE CIENTE DESCARADOS
If Me.CodCliente = 0 Or IsNull(Me.CodCliente) Then
    clie = ""
    puntosglobales = 0
    
    Me.Nombre_Provicional.Locked = False
    Me.Cuadro_combinado123 = ""
    Me.piniciales = 0
    Me.pacumulado = 0
    Me.clubguardado = 0
Else
    Dim clientebuscar As Variant
    clientebuscar = DatosClienteClubGlobal(Me.CodCliente)
    'clie = nombrecliente(Me.CodCliente.Value)
    clie = clientebuscar(1)
    'puntosglobales = PuntosGlobales2(Me.CodCliente.Value)
    puntosglobales = clientebuscar(0)
    
    Me.Nombre_Provicional.Locked = True
    Me.Cuadro_combinado123 = clie
    Me.piniciales = puntosglobales
    Me.pacumulado = IIf(Me.piniciales - Me.ppedido < 0, 0, Me.piniciales - Me.ppedido)
    Me.clubguardado = Me.CodCliente.Value
    
End If

''ENCABEZADO Y USURIO
Me.horaabierto = Now
If Me.CodCliente = 0 Or IsNull(Me.CodCliente) Then
    Me.Encabezado_automático0 = " NOTA DE PEDIDO " & Me.CodPedido
Else
    Me.Encabezado_automático0 = " NOTA DE PEDIDO " & Me.CodPedido & " - " & UCase(clie)
End If
'Dim rut As String: rut = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Diario\" & nombrelocal() & "\"
'DoCmd.TransferText acExportDelim, , "TempUltimosPedidos", rut & "UltimosPedidos.txt"

If CurrentProject.AllForms("Main Pitaya").IsLoaded Then
    [Forms]![Nota de Pedido]![colaborador] = [Forms]![Main Pitaya]![codigologin]
End If

'BOTONES ANULADO  Y MOTIVO
If Me.Anulado.Value = False Then    'Cuando el pedido esta activo
    Me.MotivoAnulado.Visible = False
    Me.EtiquetaMotivoAnulado.Visible = False
    Me.banulado.Caption = "ACTIVO"
    Me.banulado.BackColor = RGB(34, 177, 76)
    Me.Secundario57.Locked = False
Else
    Me.MotivoAnulado.Visible = True
    Me.EtiquetaMotivoAnulado.Visible = True
    Me.banulado.Caption = "ANULADO"
    Me.banulado.BackColor = RGB(237, 28, 36)
    Me.anularpedido.Visible = False
    Me.Secundario57.Locked = True
End If

'If Me.Modalidad.Value = False Then       'Cuando es Plastico
'    Me.bvidrio.Caption = "PLASTICO"
'Else
'    Me.bvidrio.Caption = "VIDRIO"
'End If


'BOTON DE TIPO DE PAGO
Select Case True

    Case Me.POS.Value = False And Me.Transferencia.Value = False And DLookup("[Comision]", "[Delivery]", "[CodDelivery]=" & Me.Delivery) = 0
        ' Caso efectivo (ninguna casilla marcada)
        Me.bpos.Caption = "EFECTIVO"
        Me.bpos.Enabled = True
        'Me.TipoDelivery.Enabled = False
        
        'If pedidocontieneproducto("Servicio Delivery", Me.CodPedido) = 0 Then ' SIn servicio de motorizado propio
            Me.tipofactura = "PEDIDO EN TIENDA"
        '    Me.CodMotorizado.Visible = False
        '    Me.Etiqueta1386.Visible = False
        'Else ' con servicio de motorizado propio
        '    Me.tipofactura = "PEDIDO CON SERVICIO DELIVERY"
            Me.CodMotorizado.Visible = True
            Me.Etiqueta1386.Visible = True
        'End If
        
        
    Case Me.POS.Value = True And Me.Transferencia.Value = True And DLookup("[Comision]", "[Delivery]", "[CodDelivery]=" & Me.Delivery) = 0
        ' Caso Transferencia
        Me.bpos.Caption = "TRANSFERENCIA"
        Me.bpos.Enabled = True
        'Me.TipoDelivery.Enabled = False
        
        'If pedidocontieneproducto("Servicio Delivery", Me.CodPedido) = 0 Then ' SIn servicio de motorizado propio
            Me.tipofactura = "PEDIDO EN TIENDA"
        '    Me.CodMotorizado.Visible = False
        '    Me.Etiqueta1386.Visible = False
        'Else ' con servicio de motorizado propio
        '    Me.tipofactura = "PEDIDO CON SERVICIO DELIVERY"
            Me.CodMotorizado.Visible = True
            Me.Etiqueta1386.Visible = True
        'End If
        
        
    Case Me.POS.Value = True And Me.Transferencia.Value = False And DLookup("[Comision]", "[Delivery]", "[CodDelivery]=" & Me.Delivery) = 0
        ' Caso POS
        Me.bpos.Caption = "POS"
        Me.bpos.Enabled = True
        'Me.TipoDelivery.Enabled = False
        
        'If pedidocontieneproducto("Servicio Delivery", Me.CodPedido) = 0 Then ' SIn servicio de motorizado propio
            Me.tipofactura = "PEDIDO EN TIENDA"
        '    Me.CodMotorizado.Visible = False
        '    Me.Etiqueta1386.Visible = False
        'Else ' con servicio de motorizado propio
        '    Me.tipofactura = "PEDIDO CON SERVICIO DELIVERY"
            Me.CodMotorizado.Visible = True
            Me.Etiqueta1386.Visible = True
        'End If
        
    Case Me.POS.Value = True And Me.Transferencia.Value = False And DLookup("[Comision]", "[Delivery]", "[CodDelivery]=" & Me.Delivery) <> 0
        ' Caso Delivery
        If Me.Delivery = 5 Then ' Hugo
            Me.bpos.Caption = "HUGO"
            Me.bpos.Enabled = False
            'Me.TipoDelivery.Enabled = True
            
            Me.tipofactura = "PEDIDO POR APLICACION HUGO"
            Me.CodMotorizado.Visible = False
            Me.Etiqueta1386.Visible = False
            
            Me.Comando1204.Visible = False
        Else 'PedidosYa
            Me.bpos.Caption = "PEDIDOSYA"
            Me.bpos.Enabled = False
            'Me.TipoDelivery.Enabled = True
            
            Me.tipofactura = "PEDIDOS YA"
            Me.CodMotorizado.Visible = False
            Me.Etiqueta1386.Visible = False
        End If
   
End Select

'Me.tipopropina = DLookup("[CodTipoPropinas]", "[TipoPropinas]", "[Valor]=" & DLookup("[Propina]", "[NotaDePedido]", "[CodPedido]=" & Me.CodPedido))
Me.tipopropina = DLookup("[Propina]", "[NotaDePedido]", "[CodPedido]=" & Me.CodPedido)

'******** Codigo de antes *********
'If Me.POS.Value = False Then 'Cuando es Efectivo
'    Me.bpos.Caption = "EFECTIVO"
'Else
'    If Me.TipoDelivery = 5 Or Me.TipoDelivery = 6 Or Me.TipoDelivery = 8 Then 'Hugo:5, piki: 6, pedidosya: 8
'        Me.bpos.Caption = "A CUENTA"
'        Me.bpos.Enabled = False
'    Else
'        Me.bpos.Caption = "TARJETA"
'    End If
'End If
'******** Codigo de antes *********


'''''''''''BLOQUEO
If codigoLocal <> 15 Then

    If Time > Me.Hora + 10 / 60 / 24 Or Date <> Me.Fecha Then ' cuando se  abre otra vez se bloquea edicion 'Pasado 1 hora de emitido se bloquea
        'MsgBox "Pedido Bloqueado,solo permitido cambiar tipo de envase y tipo de pago"
        
        Call BloquearFacturacion
        Me.pacumulado = IIf(Me.piniciales < 0, 0, Me.piniciales)
    End If
    
    If Date <> Me.Fecha Then ' bloqueo extremo cuando se abren facturas de fecha apsada
        Me.bpos.Enabled = False
        Me.Secundario57.Form.empaquetexto.Visible = False
        
    End If
    
    If Me.Impreso = True Then ' cuando ya s eimprimio
        Call BloquearFacturacion
        Me.pacumulado = IIf(Me.piniciales < 0, 0, Me.piniciales)
    End If

End If


'''''BOTONES DE MENU y DATOS DE ENVIO Y APROBACION
If Me.bpos.Caption = "PEDIDOSYA" Then
    Me.Comando68.Visible = False
Else
    Me.Comando1204.Visible = False
End If

If codigoLocal() = 15 Then
    Me.Comando1669.Visible = True 'boton menu  nrmal
    Me.botonaprobar.Visible = True
End If

Me.TipoDelivery = Me.Delivery

End Sub

Private Sub Form_Timer()
If Me.Impreso = True And codigoLocal() <> 15 And codigoLocal() <> 14 And CurrentProject.AllForms("SolicitudAnulacionPedido").IsLoaded = False Then
    If Now - Me.horaabierto > 1 / 60 / 24 Then   ' cuando se  abre otra vez se bloquea edicion
        Call Comando1658_Click
    End If
End If
End Sub

Private Sub recibecor_Exit(Cancel As Integer)
If IsNull(Me.recibecor) Then
    Me.recibecor = 0
End If
Me.cambiocor.Requery
End Sub

Private Sub recibedol_Exit(Cancel As Integer)
If IsNull(Me.recibedol) Then
    Me.recibedol = 0
End If
Me.cambiocor.Requery

End Sub

Private Sub rotulobloquear_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "desbloquearpedido"
End Sub

Private Sub TipoDelivery_Exit(Cancel As Integer)
Me.Delivery = Me.TipoDelivery
If Me.TipoDelivery = 0 Then
    Me.POS = 0
    Me.Transferencia = 0
    Me.bpos.Caption = "EFECTIVO"
    Me.TipoDelivery.Enabled = False
End If
End Sub

Private Sub tipopropina_Change()
Me.Propina = Me.tipopropina
Me.Requery
End Sub

Private Sub tipopropina_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
