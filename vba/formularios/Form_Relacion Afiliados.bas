' ==========================================================
' Modulo  : Form_Relacion Afiliados
' Tipo    : 100
' Lineas  : 8
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
DoCmd.MoveSize 5500, 500, 9000, 10000

End Sub
