' ==========================================================
' Modulo  : Form_SeleccionInternaInsumosCompraVenta
' Tipo    : 100
' Lineas  : 47
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:28
' ==========================================================
Option Compare Database



Private Sub Comando562_Click()
Me.Requery

DoCmd.OpenReport "ControlExistenciasCompraVentaSeleccionados", acViewReport

[Reports]![ControlExistenciasCompraVentaSeleccionados].Report.Printer.LeftMargin = 0
[Reports]![ControlExistenciasCompraVentaSeleccionados].Report.Printer.RightMargin = 0
[Reports]![ControlExistenciasCompraVentaSeleccionados].Report.Printer.BottomMargin = 0
[Reports]![ControlExistenciasCompraVentaSeleccionados].Report.Printer.TopMargin = 0

DoCmd.PrintOut , 1, 1
DoCmd.Close acReport, "ControlExistenciasCompraVentaSeleccionados"

Call importartablaespecifica("Main_DB", "DBIngredientes", "DBIngredientes", 1)
DoCmd.Close acForm, "SeleccionInternaInsumosCompraventa"

End Sub

Private Sub Comando790_Click()
'If Me.Comando790.Caption = "QUITAR SELECCION" Then
'    Me.Comando790.Caption = "SELECCIONAR TODO"
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL "UPDATE DBBatidos SET DBBatidos.CodGrupo = 0 WHERE DBBatidos.CodGrupo = 7 AND CotiPrincipalProdCompraVenta(DBBatidos.CodBatido) = " & Me.coti
'    DoCmd.SetWarnings True
'    'Me.Requery
'Else
'    Me.Comando790.Caption = "QUITAR SELECCION"
'    Call importartablaespecifica("Main_DB", "DBIngredientes", "DBIngredientes", 1)
'    Me.Requery
'End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub Verificación575_Click()
'DoCmd.SetWarnings False
'DoCmd.RunSQL "UPDATE DBBatidos SET DBBatidos.CodGrupo = 0 WHERE DBBatidos.CodGrupo = 7 AND CotiPrincipalProdCompraVenta(DBBatidos.CodBatido) = " & Me.coti
'DoCmd.SetWarnings True
'MsgBox ""
End Sub
