' ==========================================================
' Modulo  : Form_Consumo Ingrediente Otros
' Tipo    : 100
' Lineas  : 14
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando15_Click()
Me.Requery
End Sub

Private Sub Comando342_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
