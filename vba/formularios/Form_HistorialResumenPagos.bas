' ==========================================================
' Modulo  : Form_HistorialResumenPagos
' Tipo    : 100  |  Lineas: 48
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database



Private Sub Comando148_Click()
DoCmd.OpenForm "Resumen Pagos Proovedor"
[Forms]![Resumen Pagos Proovedor]![aresumen] = Me.codresumenpago
[Forms]![Resumen Pagos Proovedor]![CodigoBusqueda] = Me.CodProovedor
[Forms]![Resumen Pagos Proovedor]![sucursalbusqueda] = Me.Sucursal
[Forms]![Resumen Pagos Proovedor].Form.Requery
End Sub

Private Sub Comando209_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO resumenpago(CodProovedor, fecha)" & _
" values (" & Me.CodigoBusqueda & ", #" & Date & "#)"
DoCmd.SetWarnings True

Me.Requery

Dim codresumen As Long
codresumen = ultimoresumen()

DoCmd.OpenForm "Resumen Pagos Proovedor"

[Forms]![Resumen Pagos Proovedor]![aresumen] = codresumen
[Forms]![Resumen Pagos Proovedor]![CodigoBusqueda] = Me.CodigoBusqueda
[Forms]![Resumen Pagos Proovedor].Form.Requery
End Sub


Private Sub Comando253_Click()
If Me.Comando253.Caption = "FILTRAR" Then
    Me.Filter = "[codproovedor] = " & Me.CodigoBusqueda
    Me.FilterOn = True
    Me.Comando253.Caption = "VER TODOS"
Else
    Me.FilterOn = False
    Me.Comando253.Caption = "FILTRAR"
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"

End Sub


