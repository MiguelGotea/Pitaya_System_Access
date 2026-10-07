' ==========================================================
' Modulo  : Form_Cierre por Turno 2
' Tipo    : 100
' Lineas  : 19
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database







Private Sub Comando426_Click()
DoCmd.Close acForm, "Cierre por Turno"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False


End Sub

