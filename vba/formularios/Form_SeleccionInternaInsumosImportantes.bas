' ==========================================================
' Modulo  : Form_SeleccionInternaInsumosImportantes
' Tipo    : 100  |  Lineas: 45
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:18
' ==========================================================

Option Compare Database






Public Sub Comando562_Click()
Me.Requery

DoCmd.OpenReport "ControlInsumosImportantesSeleccionados", acViewReport

[Reports]![ControlInsumosImportantesSeleccionados].Report.Printer.LeftMargin = 0
[Reports]![ControlInsumosImportantesSeleccionados].Report.Printer.RightMargin = 0
[Reports]![ControlInsumosImportantesSeleccionados].Report.Printer.BottomMargin = 0
[Reports]![ControlInsumosImportantesSeleccionados].Report.Printer.TopMargin = 0

DoCmd.PrintOut , 1, 1
DoCmd.Close acReport, "ControlInsumosImportantesSeleccionados"

Call importartablaespecifica("Main_DB", "DBIngredientes", "DBIngredientes", 1)
DoCmd.OpenForm "SeleccionInternaInsumosCompraventa"
DoCmd.Close acForm, "SeleccionInternaInsumosImportantes"

End Sub



Private Sub Comando790_Click()
If Me.Comando790.Caption = "QUITAR SELECCION" Then
    Me.Comando790.Caption = "SELECCIONAR TODO"
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE DBIngredientes SET DBIngredientes.ListaImportante = 0"
    DoCmd.SetWarnings True
    'Me.Requery
Else
    Me.Comando790.Caption = "QUITAR SELECCION"
    Call importartablaespecifica("Main_DB", "DBIngredientes", "DBIngredientes", 1)
    Me.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
