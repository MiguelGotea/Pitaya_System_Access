' ==========================================================
' Modulo  : Form_Conv_Sem_Coti
' Tipo    : 100  |  Lineas: 6
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
