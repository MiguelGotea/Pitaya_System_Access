' ==========================================================
' Modulo  : Form_RegistroVacaciones
' Tipo    : 100  |  Lineas: 7
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:01
' ==========================================================

Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
