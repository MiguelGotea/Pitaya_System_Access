' ==========================================================
' Modulo  : Form_Menu PITAYA Delivery
' Tipo    : 100
' Lineas  : 52
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================


Private Sub Comando538_Click()
On Error Resume Next
''Version cerradno y abir nuevamente
'Dim codi As Long
'Dim men As String
'men = Me.Name ' nombre del formulario
'codi = [Forms]![Nota de Pedido]![CodPedido]
'DoCmd.Close acForm, "Nota de Pedido"
'DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & codi
'DoCmd.Close acForm, men

'Version Fast
If Me.solovista = 0 Then
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![Nota de Pedido].Form.Secundario57.Requery
    [Forms]![Nota de Pedido].Form.ppedido.Requery
    [Forms]![Nota de Pedido].Form.Texto193.Requery
    [Forms]![Nota de Pedido].Form.Texto1278.Requery

    [Forms]![Nota de Pedido].Form.cambiocor.Requery
    [Forms]![Nota de Pedido].Form.recibecor.SetFocus
End If
[Forms]![Nota de Pedido].Form.Secundario57.Requery
DoCmd.Close acForm, "Menu PITAYA Delivery"
End Sub













Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False



End Sub

Private Sub salirsinguardar_Click()
DoCmd.Close
End Sub

