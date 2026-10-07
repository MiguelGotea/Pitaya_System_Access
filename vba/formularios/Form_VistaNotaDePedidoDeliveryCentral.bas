' ==========================================================
' Modulo  : Form_VistaNotaDePedidoDeliveryCentral
' Tipo    : 100
' Lineas  : 129
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database





Private Sub Comando155_Click()
If MsgBox("Desea cargar pedido a la facturacion de la sucursal?", vbYesNo, "CONFIRMACION") <> vbYes Then
    Exit Sub
End If

Dim nuevopedi As Long
Dim pedix As Long

pedix = Me.CodPedido
nuevopedi = ultimopedidofacturado() + 1

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO statusPedidosCentral(CodPedidoCentral, HoraAprobadoCentral," & _
" AprobadoCentral, Sucursal, Fecha, CodPedidoSucursal,HoraAprobadoSucursal,AprobadoSucursal)" & _
" values (" & pedix & ", #" & DLookup("[HoraAprobadoCentral]", "[StatusPedidosCentralDeliveryCentral]", "[CodPedidoCentral]=" & pedix) & "#," & _
" -1," & codigoLocal() & ", Date, " & nuevopedi & ", #" & Now() & "#, -1)"
DoCmd.SetWarnings True

If IsNull(Me.CodMotorizado) Or Me.CodMotorizado = "" Then
    Call moverfacturacentralasucursal(pedix, nuevopedi, 0)
Else
    Call moverfacturacentralasucursal(pedix, nuevopedi, 1)
End If

Call moverdetallefacturacentralasucursal(pedix, nuevopedi)

Call EnviarMensajeTelegram("Pedido Facturado " & pedix & vbCrLf & _
    "Pedido en sucursal: " & nuevopedi, grupotpedidospitayacentral())

MsgBox "Pedido " & nuevopedi & " creado", vbCritical

'DoCmd.Close acForm, "HistorialVentasDeliveryCentral"
[Forms]![Main Pitaya].Form.Requery
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & nuevopedi

Call Comando1658_Click
End Sub

Private Sub Comando1658_Click()

DoCmd.Close acForm, "VistaNotaDePedidoDeliveryCentral"
End Sub




Private Sub Form_Open(Cancel As Integer)
'On Error Resume Next
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False



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

        'End If
        
    Case Me.POS.Value = True And Me.Transferencia.Value = False And DLookup("[Comision]", "[Delivery]", "[CodDelivery]=" & Me.Delivery) <> 0
        ' Caso Delivery
        If Me.Delivery = 5 Then ' Hugo
            Me.bpos.Caption = "HUGO"
            Me.bpos.Enabled = False
            'Me.TipoDelivery.Enabled = True
            
            Me.tipofactura = "PEDIDO POR APLICACION HUGO"

        Else 'PedidosYa
            Me.bpos.Caption = "PEDIDOSYA"
            Me.bpos.Enabled = False
            'Me.TipoDelivery.Enabled = True
            
            Me.tipofactura = "PEDIDOS YA"

        End If
   
End Select


End Sub

