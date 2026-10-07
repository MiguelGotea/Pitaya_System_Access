' ==========================================================
' Modulo  : Form_Relacion Pedidos Dia
' Tipo    : 100  |  Lineas: 20
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database

Private Sub Comando119_Click()
Me.Filter = "Fecha = #" & Me.afecha & "#"
Me.FilterOn = True
Me.Requery
End Sub



Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.afecha = Date
End Sub
