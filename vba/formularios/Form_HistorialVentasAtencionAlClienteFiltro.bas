' ==========================================================
' Modulo  : Form_HistorialVentasAtencionAlClienteFiltro
' Tipo    : 100  |  Lineas: 21
' Proyecto: Database3
' Exportado: 2026-10-14 21:56:11
' ==========================================================

Option Compare Database



Private Sub Comando102_Click()

Me.Requery

End Sub

Private Sub Comando107_Click()
DoCmd.OpenForm "Nota de Pedido", , , "[CodPedido]=" & Me.CodPedido
End Sub



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
