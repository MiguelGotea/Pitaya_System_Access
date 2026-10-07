' ==========================================================
' Modulo  : Form_HistorialVentasDeliveryCentral
' Tipo    : 100  |  Lineas: 36
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database



Private Sub Form_Close()
Me.RecordSource = ""
DoCmd.RunSQL "DROP TABLE NotaDePedidoDeliveryCentral"
DoCmd.RunSQL "DROP TABLE SubPedidoDeliveryCentral"
DoCmd.RunSQL "DROP TABLE StatusPedidosCentralDeliveryCentral"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub



Private Sub Comando50_Click()
If Me.Estado = "Anulado" Then
    MsgBox "No se pueden generar pedidos anulados desde la central"
    Exit Sub
End If

If Me.pedidosucursal = 0 Then
    DoCmd.OpenForm "VistaNotaDePedidoDeliveryCentral", , , "[CodPedido]=" & Me.CodPedido
Else
    MsgBox "Pedido ya se paso a la facturacion de la sucursal"
End If
End Sub




