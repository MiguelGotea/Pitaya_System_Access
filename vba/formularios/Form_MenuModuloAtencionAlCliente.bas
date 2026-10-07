' ==========================================================
' Modulo  : Form_MenuModuloAtencionAlCliente
' Tipo    : 100  |  Lineas: 34
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database


Private Sub Comando280_Click()
Call actualizardatosclubmixedglobal
'Call DescargarTablaCompleta("clientesclub", "clientesclubexterno", "sucursal <> " & codigolocal())
MsgBox "Tablas de clientes actualizadas"
End Sub



Private Sub Comando311_Click()
DoCmd.OpenForm "Clientes Club Pitaya"
End Sub


Private Sub Comando348_Click()
DoCmd.OpenForm "PromocionesVigentes"

End Sub

Private Sub Comando371_Click()
Call eliminartablasmain
Call importartablasmain
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub Comando285_Click()
DoCmd.Quit
End Sub
