' ==========================================================
' Modulo  : Form_SolicitudAnulacionPedido
' Tipo    : 100  |  Lineas: 115
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database

Private Sub botonbloquear_Click()
'On Error GoTo Nulo

If Me.Texto1305.Caption = "SOLICITUD DE ANULACION DE PEDIDO" Then 'validacion de inputs llenados corrtamente

    If IsNull(Me.pedidocambio) Or Me.pedidocambio = "" Then
        MsgBox "Ingresar el numero de pedido que se hiso ya corregido"
        Exit Sub
    End If
    
    If CLng(Me.pedidoanular) = CLng(Me.pedidocambio) Then
        MsgBox "El numero de pedido corregido no puede ser el mismo numero del que se esta anulando"
        Exit Sub
    End If
    
    If IsNull(DLookup("[CodPedido]", "[NotaDePedido]", "[CodPedido]=" & CLng(Me.pedidocambio))) Then
        MsgBox "Ingresar un numero de pedido corregido que ya haya sido creado e impreso"
        Exit Sub
    End If

End If

If IsNull(Me.pedidomotivo) Or Me.pedidomotivo = "" Then ' validacion global de inputs
    MsgBox "Debe ingresar el sustento por el que se esta anulando el pedido"
    Exit Sub
End If

If IsNull(Me.pedidotipoanulacion) Or Me.pedidotipoanulacion = "" Then ' validacion global de inputs
    MsgBox "Ingresar el tipo de motivo de anulacion del pedido"
    Exit Sub
End If


If Me.Texto1305.Caption = "SOLICITUD DE ANULACION DE PEDIDO" Then 'cuando es una solicitud web/telegram

    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO AnulacionPedidos(CodPedido, HoraSolicitada, Modalidad, CodPedidoCambio, Motivo, CodMotivoAnulacion)" & _
                 " values (" & Me.pedidoanular & ", #" & Time & "#, 2, " & Me.pedidocambio & ", '" & Me.pedidomotivo & "', " & Me.pedidotipoanulacion & ")"
    DoCmd.SetWarnings True
    
    If EsSistemaDeTienda() Then

        Dim mensajeg As String
        'mensajeg = "Solicitud de Anulacion de Pedido: " & Me.pedidoanular & " " & statusnotaddepedido(Me.pedidoanular) & " (" & DLookup("[TotalGuardado]", "[NotaDePedido]", "[CodPedido]=" & Me.pedidoanular) & ")" & vbCrLf & vbCrLf & _
                  detalletextopedido(Me.pedidoanular) & vbCrLf & _
                  "Solicitado por: " & Me.pedidooperario & vbCrLf & _
                  "Motivo: " & Me.pedidomotivo & vbCrLf & vbCrLf & _
                  "Pedido generado como correccion: " & Me.pedidocambio & " " & statusnotaddepedido(Me.pedidocambio) & " (" & DLookup("[TotalGuardado]", "[NotaDePedido]", "[CodPedido]=" & Me.pedidocambio) & ")" & vbCrLf & vbCrLf & detalletextopedido(CLng(Me.pedidocambio))
        
        'Call EnviarMensajeTelegram(mensajeg, grupotanulaciones())
        
        Call SyncVentasPedido(Me.pedidocambio) ' enviar dato individual de pedido
        Call SyncVentasPedido(Me.pedidoanular) ' enviar dato individual de pedido
        
        Call SyncEnviarAnulacionesPendientes

        MsgBox "Solicitud de anulacion enviada correctamente, el pedido se anulara en el transcurso del turno, los puntos de club se restaran una vez anulado el pedido"
    
    End If
     
Else ' anulado automaticamente sin pasar por solicitud
                        
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO AnulacionPedidos(CodPedido, HoraSolicitada, HoraAnulada, Status, Modalidad, CodPedidoCambio, Motivo, CodMotivoAnulacion)" & _
                 " values (" & Me.pedidoanular & ", #" & Time & "#, #" & Time & "#,1,1, 0, '" & Me.pedidomotivo & "', " & Me.pedidotipoanulacion & ")"
    DoCmd.SetWarnings True
    
    If codigoLocal() = 15 Then 'cuando se anula un pedido desde delivery central
        If statusaprobaciondeliverycentral(Me.pedidoanular) = True Then 'si hay registro de pedido mandado de central a sucursal
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * FROM StatusPedidosCentral WHERE CodPedidoCentral = " & Me.pedidoanular
            DoCmd.SetWarnings True
        End If
    End If
                     
    [Forms]![Nota de Pedido]![Anulado].Value = True
    [Forms]![Nota de Pedido]![MotivoAnulado].Visible = True
    [Forms]![Nota de Pedido]![EtiquetaMotivoAnulado].Visible = True
    [Forms]![Nota de Pedido]![banulado].Caption = "ANULADO"
    [Forms]![Nota de Pedido]![banulado].BackColor = RGB(237, 28, 36)
    
    [Forms]![Nota de Pedido]![MotivoAnulado] = Me.pedidomotivo
    
    
    
    'Call Forms("[Nota de Pedido]").BloquearFacturacion
    MsgBox "Pedido anulado, este pedido pasara a revision y auditoria para su verificacion y confirmacion"
    
    'Call EnviarMensajeTelegram("Anulando Pedido " & Me.pedidoanular & " Usuario " & Me.pedidooperario & " Motivo " & Me.pedidomotivo, grupotanulaciones())
End If

Call Forms("Nota de Pedido").BloquearFacturacion
DoCmd.Close acForm, "SolicitudAnulacionPedido"
Exit Sub
Nulo:
MsgBox "Solicitar anulacion directamente, problemas con el servidor"

End Sub

Private Sub Comando1253_Click()
DoCmd.Close acForm, "SolicitudAnulacionPedido"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub pedidotipoanulacion_Exit(Cancel As Integer)
If Me.pedidotipoanulacion = 6 Then ' Otros
    MsgBox "ingresar cual es el sustento por el que se esta anulando este pedido"
End If
End Sub
