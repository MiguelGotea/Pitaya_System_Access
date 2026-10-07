' ==========================================================
' Modulo  : Form_GenerarIngresoDeCompras
' Tipo    : 100  |  Lineas: 53
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database


Private Sub Comando112_Click()
If fechaac.Value < Date Then
fechaac.Value = fechaac.Value + 1
Me.Requery
End If
End Sub

Private Sub Comando113_Click()
fechaac.Value = fechaac.Value - 1
If fechaac.Value < Date - 7 And codigoLocal() <> 0 Then
    fechaac.Value = fechaac.Value + 1
End If
Me.Requery
End Sub

Private Sub Comando434_Click()
DoCmd.OpenForm "AutogenerarIngresodeCompras", , , "[NumeroFactura]='" & Me.NumeroFactura & "'"
End Sub

Private Sub fechaac_Change()
Me.Requery
End Sub

Private Sub Form_Close()
If CurrentProject.AllForms("Contador Caja").IsLoaded Then
    [Forms]![Contador Caja]![productoscomprassistema].height = 280 * CantidadProductosCompradosCajaDia(Date)
    [Forms]![Contador Caja].montocomprassistema.Requery
    [Forms]![Contador Caja].vvales.Requery
    [Forms]![Contador Caja].cantidadcomprassistema.Requery
    [Forms]![Contador Caja].productoscomprassistema.Requery
    [Forms]![Contador Caja].Form.Requery
End If

If CurrentProject.AllForms("Contador Caja 2").IsLoaded Then
    [Forms]![Contador Caja 2]![productoscomprassistema].height = 280 * CantidadProductosCompradosCajaDia(Date)
    [Forms]![Contador Caja 2].montocomprassistema.Requery
    [Forms]![Contador Caja 2].vvales.Requery
    [Forms]![Contador Caja 2].cantidadcomprassistema.Requery
    [Forms]![Contador Caja 2].productoscomprassistema.Requery
    [Forms]![Contador Caja 2].Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

Me.Requery
End Sub


