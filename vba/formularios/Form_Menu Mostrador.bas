' ==========================================================
' Modulo  : Form_Menu Mostrador
' Tipo    : 100  |  Lineas: 12
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================



Private Sub Comando538_Click()
DoCmd.Close

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

