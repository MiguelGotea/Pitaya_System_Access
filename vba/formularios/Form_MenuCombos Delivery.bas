' ==========================================================
' Modulo  : Form_MenuCombos Delivery
' Tipo    : 100
' Lineas  : 7
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:30
' ==========================================================
Private Sub Comando538_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
