' ==========================================================
' Modulo  : Form_Calculadora Mezclas
' Tipo    : 100  |  Lineas: 11
' Proyecto: Database3
' Exportado: 2026-10-22 22:23:01
' ==========================================================

Option Compare Database

Private Sub Comando39_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

