' ==========================================================
' Modulo  : Form_Facturas Por Pagar
' Tipo    : 100  |  Lineas: 9
' Proyecto: Database3
' Exportado: 2026-10-14 21:56:07
' ==========================================================

Option Compare Database



Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub
