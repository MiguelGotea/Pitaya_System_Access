' ==========================================================
' Modulo  : Form_EventosDia
' Tipo    : 100  |  Lineas: 6
' Proyecto: Database3
' Exportado: 2026-10-14 21:56:07
' ==========================================================

Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
