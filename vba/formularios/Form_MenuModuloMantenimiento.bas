' ==========================================================
' Modulo  : Form_MenuModuloMantenimiento
' Tipo    : 100
' Lineas  : 42
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:29
' ==========================================================
Option Compare Database


Private Sub Comando1217_Click()
DoCmd.Quit
End Sub




Private Sub Comando1674_Click()
Call eliminartablasmain

Call importartablasmain

MsgBox "Tablas Main Actualizadas"
End Sub


Private Sub Comando1702_Click()
Call descargardatoscentralcopia("mantenimiento")
MsgBox "Datos de la central completos"
End Sub



Private Sub Comando2048_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "Mantenimiento"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "BATIDOS PITAYA 0"
Me.ShortcutMenu = False

End Sub





