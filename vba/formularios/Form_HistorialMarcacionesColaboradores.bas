' ==========================================================
' Modulo  : Form_HistorialMarcacionesColaboradores
' Tipo    : 100
' Lineas  : 13
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database





Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub


