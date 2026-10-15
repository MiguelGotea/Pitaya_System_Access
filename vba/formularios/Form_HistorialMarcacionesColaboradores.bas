' ==========================================================
' Modulo  : Form_HistorialMarcacionesColaboradores
' Tipo    : 100  |  Lineas: 13
' Proyecto: Database3
' Exportado: 2026-10-14 21:56:11
' ==========================================================

Option Compare Database





Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub


