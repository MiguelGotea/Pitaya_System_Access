' ==========================================================
' Modulo  : Form_MenuPitayaProduccion0
' Tipo    : 100  |  Lineas: 78
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub Comando1217_Click()
DoCmd.Quit
End Sub

Private Sub Comando1521_Click()
DoCmd.OpenForm "DesempenoOperarioPorcionamiento"
End Sub

Private Sub Comando1524_Click()
DoCmd.OpenForm "DesempenoOperarioProcesamiento"
End Sub

Private Sub Comando1530_Click()
DoCmd.OpenForm "RequerimientoPorcionesSemana"
End Sub



Private Sub Comando1559_Click()
DoCmd.OpenForm "PlanProduccionMarcaPitaya"
End Sub



Private Sub Comando1570_Click()
DoCmd.OpenForm "PlanPorcionesSemanaTotal"
End Sub

Private Sub Comando1576_Click()
Call eliminartablasmain

Call importartablasmain

Call importartablasweb
MsgBox "Tablas Main Actualizadas"
End Sub

Private Sub Comando1621_Click()
DoCmd.OpenForm "Calculo Conversion Cotizacion Semana"
End Sub





Private Sub Comando1666_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "Produccion"
End Sub

Private Sub Comando1702_Click()
Call descargardatoscentralcopia("produccion")
MsgBox "Datos de la central completos"

End Sub

Private Sub Comando1490_Click()
DoCmd.OpenForm "PlanProduccionMarcaPitaya"
[Forms]![PlanProduccionMarcaPitaya].asemana = [Forms]![MenuPitayaProduccion0].semanaplanproduccion
Call Forms("[PlanProduccionMarcaPitaya]").asemana_Exit(0)
Call Forms("[PlanProduccionMarcaPitaya]").Comando170_Click

DoCmd.OpenForm "PlanPorcionesSemanaTotal"
[Forms]![PlanPorcionesSemanaTotal].asemana = [Forms]![MenuPitayaProduccion0].semanaplanproduccion
Call Forms("[PlanPorcionesSemanaTotal]").asemana_Exit(0)
Call Forms("[PlanPorcionesSemanaTotal]").Comando1040_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "BATIDOS PITAYA 0"
Me.ShortcutMenu = False

End Sub


