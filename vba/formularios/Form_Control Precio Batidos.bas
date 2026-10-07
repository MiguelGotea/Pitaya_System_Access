' ==========================================================
' Modulo  : Form_Control Precio Batidos
' Tipo    : 100
' Lineas  : 10
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database

Private Sub Comando41_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
