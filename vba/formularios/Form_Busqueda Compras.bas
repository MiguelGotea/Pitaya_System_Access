' ==========================================================
' Modulo  : Form_Busqueda Compras
' Tipo    : 100  |  Lineas: 11
' Proyecto: Database3
' Exportado: 2026-10-22 22:22:59
' ==========================================================

Option Compare Database

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.Requery
End Sub
