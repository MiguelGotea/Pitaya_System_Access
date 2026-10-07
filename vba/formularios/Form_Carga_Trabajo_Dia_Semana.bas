' ==========================================================
' Modulo  : Form_Carga_Trabajo_Dia_Semana
' Tipo    : 100
' Lineas  : 7
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub
