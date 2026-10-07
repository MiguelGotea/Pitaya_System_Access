' ==========================================================
' Modulo  : Form_Menu Adicionales
' Tipo    : 100
' Lineas  : 13
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:24
' ==========================================================
Private Sub Comando538_Click()
[Forms]![Nota de Pedido].Form.Secundario57.Requery

[Forms]![Nota de Pedido].Form.Texto1278.Requery

[Forms]![Nota de Pedido].Form.recibecor.SetFocus
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub
