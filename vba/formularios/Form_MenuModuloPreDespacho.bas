' ==========================================================
' Modulo  : Form_MenuModuloPreDespacho
' Tipo    : 100  |  Lineas: 51
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
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
Call descargardatoscentralcopia("predespacho")
MsgBox "Datos de la central completos"
End Sub



Private Sub Comando2048_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "PreDespacho"
End Sub

Private Sub Comando774_Click()
'Modulo Despacho
Call importartablaespecifica("Despacho", "PreIngresoPitaya", "PreIngresoPitaya", 1)
Call importartablaespecifica("Despacho", "SubPreIngresosPitaya", "SubPreIngresosPitaya", 1)

DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").modopredespacho
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "BATIDOS PITAYA 0"
Me.ShortcutMenu = False

End Sub





