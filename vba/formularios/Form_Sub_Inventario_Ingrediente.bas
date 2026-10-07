' ==========================================================
' Modulo  : Form_Sub_Inventario_Ingrediente
' Tipo    : 100  |  Lineas: 15
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

'If comprobarInventario() = #1/1/2000# Then
'No hay inventario temporal
'Else
 '   If MsgBox("Existe un inventario en proceso de registro, ¿Desea abrirlo?", vbYesNo) = vbYes Then
    
  '  [Forms]![LogueoUsuario]![CodigoBusqueda].Enabled = False

End Sub
