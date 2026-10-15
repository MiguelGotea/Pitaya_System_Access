' ==========================================================
' Modulo  : Form_Alarmas Programadas
' Tipo    : 100  |  Lineas: 7
' Proyecto: Database3
' Exportado: 2026-10-14 21:56:08
' ==========================================================

Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Call multiplebeep(10)
End Sub
