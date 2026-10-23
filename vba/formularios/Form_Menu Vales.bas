' ==========================================================
' Modulo  : Form_Menu Vales
' Tipo    : 100  |  Lineas: 10
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:08
' ==========================================================



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub

Private Sub regresar_Click()
DoCmd.Close
End Sub
