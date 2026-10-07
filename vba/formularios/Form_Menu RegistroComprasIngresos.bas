' ==========================================================
' Modulo  : Form_Menu RegistroComprasIngresos
' Tipo    : 100  |  Lineas: 107
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database




Private Sub Comando1702_Click()
Call descargardatoscentralcopia("compras")
MsgBox "Datos de la central completos"
End Sub

Private Sub Comando2058_Click()
DoCmd.OpenForm "CompraInsumosFijosConsumiblesGlobal"
End Sub

Private Sub Comando209_Click()
DoCmd.OpenForm "CompraProcesamientoSemana"
End Sub

Private Sub Comando246_Click()
DoCmd.Quit
End Sub

Private Sub Comando254_Click()
DoCmd.OpenForm "HistorialOrdenDeCompra"
End Sub

Private Sub Comando255_Click()
DoCmd.OpenForm "HistorialResumenPagos"
End Sub

Private Sub Comando269_Click()
Call eliminartablasmain

Call importartablasmain


MsgBox "Tablas Main Actualizadas"
End Sub



Private Sub Comando285_Click()
DoCmd.OpenForm "ComprasSemanalesRefrigerados"
End Sub

Private Sub Comando361_Click()
DoCmd.OpenForm "ComprasSemanalesSecos"
End Sub



Private Sub Comando385_Click()
DoCmd.OpenForm "Lista de Proovedores"
End Sub

Private Sub Comando396_Click()
DoCmd.OpenForm "EstandarInsumosFijosConsumiblesGlobal"
End Sub

Private Sub Comando401_Click()
DoCmd.OpenForm "HistorialCostoUnitarioIngredientes"
End Sub





Private Sub Comando413_Click()
DoCmd.SetWarnings False
DoCmd.CopyObject , "ComprasSucursal", acTable, "Compras"
DoCmd.RunSQL "ALTER TABLE ComprasSucursal ADD COLUMN Sucursal Integer, Elegir YesNo"
DoCmd.RunSQL "DELETE * FROM ComprasSucursal"
DoCmd.SetWarnings True
DoCmd.OpenForm "ComprasCentralSucursal"
End Sub

Private Sub Comando423_Click()
DoCmd.OpenForm "IngresoClavePrivado"
[Forms]![IngresoClavePrivado].Direccion = "actualizartablasmodulo"
[Forms]![IngresoClavePrivado].variable = "Compras"
End Sub

Private Sub Comando85_Click()
DoCmd.OpenForm "CompraxProovedor"
End Sub

Private Sub Comando372_Click()
DoCmd.OpenForm "CompraProcesamientoSemana"
[Forms]![CompraProcesamientoSemana].asemana = [Forms]![Menu RegistroComprasIngresos].semanaplancompra
Call Forms("[CompraProcesamientoSemana]").asemana_Exit(0)
Call Forms("[CompraProcesamientoSemana]").Comando199_Click

DoCmd.OpenForm "ComprasSemanalesRefrigerados"
[Forms]![ComprasSemanalesRefrigerados].asemana = [Forms]![Menu RegistroComprasIngresos].semanaplancompra
Call Forms("[ComprasSemanalesRefrigerados]").asemana_Exit(0)
Call Forms("[ComprasSemanalesRefrigerados]").Comando170_Click

DoCmd.OpenForm "ComprasSemanalesSecos"
[Forms]![ComprasSemanalesSecos].asemana = [Forms]![Menu RegistroComprasIngresos].semanaplancompra
Call Forms("[ComprasSemanalesSecos]").asemana_Exit(0)
Call Forms("[ComprasSemanalesSecos]").Comando170_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " ¦¦"
Me.ShortcutMenu = False
End Sub
