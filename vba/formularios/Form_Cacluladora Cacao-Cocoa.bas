' ==========================================================
' Modulo  : Form_Cacluladora Cacao-Cocoa
' Tipo    : 100  |  Lineas: 12
' Proyecto: Database3
' Exportado: 2026-10-14 21:56:07
' ==========================================================

Option Compare Database

Private Sub Comando39_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

