' ==========================================================
' Modulo  : Form_Historial Ventas
' Tipo    : 100
' Lineas  : 30
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================
Option Compare Database

Private Sub agrupo_Change()
Me.asubgrupo.Value = Null
Me.aproducto.Value = Null
End Sub

Private Sub agrupo_Exit(Cancel As Integer)
Me.asubgrupo.Requery
Me.aproducto.Requery
End Sub

Private Sub asubgrupo_Change()
Me.aproducto.Value = Null
End Sub

Private Sub asubgrupo_Exit(Cancel As Integer)
Me.aproducto.Requery
End Sub



Private Sub Comando56_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
