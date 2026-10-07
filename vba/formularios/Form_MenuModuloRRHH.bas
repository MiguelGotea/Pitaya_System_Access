' ==========================================================
' Modulo  : Form_MenuModuloRRHH
' Tipo    : 100  |  Lineas: 45
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database
















Private Sub Comando311_Click()
DoCmd.OpenForm "RelacionOperariosGlobal"
End Sub





Private Sub Comando383_Click()
Call eliminartablasmain

Call importartablasmain

MsgBox "Tablas Main Actualizadas"
End Sub





Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub Comando285_Click()
DoCmd.Quit
End Sub
