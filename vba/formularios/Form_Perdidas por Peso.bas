' ==========================================================
' Modulo  : Form_Perdidas por Peso
' Tipo    : 100  |  Lineas: 20
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando100_Click()
Me.Requery
End Sub

Private Sub Comando103_Click()
Me.afecha.Value = Me.afecha.Value - 1
Me.Requery
End Sub

Private Sub Comando92_Click()
Me.afecha.Value = Me.afecha.Value + 1
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
