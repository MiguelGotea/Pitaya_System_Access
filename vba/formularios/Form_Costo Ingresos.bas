' ==========================================================
' Modulo  : Form_Costo Ingresos
' Tipo    : 100  |  Lineas: 22
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:11
' ==========================================================

Option Compare Database



Private Sub Comando458_Click()
Me.Requery
End Sub

Private Sub Comando507_Click()
Me.semana1 = Me.semana1 - 1
Me.Requery
End Sub

Private Sub Comando511_Click()
Me.semana1 = Me.semana1 + 1
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
