' ==========================================================
' Modulo  : Form_MenuModuloSupervisionSucursales
' Tipo    : 100
' Lineas  : 30
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database

Private Sub Comando311_Click()
Call importartablasweb
End Sub

Private Sub Comando373_Click()
DoCmd.OpenForm "HorariosAprobadosOperariosSemana"
[Forms]![HorariosAprobadosOperariosSemana]![Comando1015].Visible = True
End Sub

Private Sub Comando376_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "SupervisionSucursales"
End Sub

Private Sub Comando85_Click()
DoCmd.OpenForm "AprobacionHorariosOperariosSemana"

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub Comando285_Click()
DoCmd.Quit
End Sub
