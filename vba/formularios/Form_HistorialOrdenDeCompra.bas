' ==========================================================
' Modulo  : Form_HistorialOrdenDeCompra
' Tipo    : 100  |  Lineas: 101
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:14
' ==========================================================

Option Compare Database



Private Sub CodigoBusqueda_Exit(Cancel As Integer)
If Me.Comando222.Caption = "VER TODOS" Then
    Me.FilterOn = False
    Me.Comando222.Caption = "FILTRAR"
End If
End Sub

Private Sub Comando148_Click()
DoCmd.OpenForm "OrdenDeCompraPlantilla"

[Forms]![OrdenDeCompraPlantilla]![ordenbusqueda] = Me.codordendecompra
[Forms]![OrdenDeCompraPlantilla]![afechaorden] = Me.fechaorden
[Forms]![OrdenDeCompraPlantilla]![aproovedor] = Me.CodProovedor
[Forms]![OrdenDeCompraPlantilla]![tipopago] = Me.tipopago

If Me.resumendepago <> 0 Then
    [Forms]![OrdenDeCompraPlantilla]![Encabezado_automático0].Caption = "RESUMEN DE PAGO"
    [Forms]![OrdenDeCompraPlantilla]![Comando820].Visible = False
    [Forms]![OrdenDeCompraPlantilla]![Comando106].Visible = False
    [Forms]![OrdenDeCompraPlantilla]![costounitario].Enabled = False
    [Forms]![OrdenDeCompraPlantilla]![cantidadorden].Enabled = False
    [Forms]![OrdenDeCompraPlantilla]![aplicaiva].Enabled = False
End If

[Forms]![OrdenDeCompraPlantilla].Requery
Call Forms("[OrdenDeCompraPlantilla]").actualizarmodopagoinfo
End Sub

Private Sub Comando187_Click()
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO OrdenDeCompra(codproovedor, fechaorden)" & _
" values (" & Me.CodigoBusqueda & ", #" & Date & "#)"
DoCmd.SetWarnings True

Me.Requery

Dim codordencompra As Long
codordencompra = ultimaordendecompra()

DoCmd.OpenForm "OrdenDeCompraPlantilla"

[Forms]![OrdenDeCompraPlantilla]![ordenbusqueda] = codordencompra
[Forms]![OrdenDeCompraPlantilla]![afechaorden] = Date
[Forms]![OrdenDeCompraPlantilla]![aproovedor] = Me.CodigoBusqueda
[Forms]![OrdenDeCompraPlantilla].Requery
End Sub



Private Sub Comando222_Click()
If Me.Comando222.Caption = "FILTRAR" Then
    Me.Filter = "[codproovedor] = " & Me.CodigoBusqueda
    Me.FilterOn = True
    Me.Comando222.Caption = "VER TODOS"
Else
    Me.FilterOn = False
    Me.Comando222.Caption = "FILTRAR"
End If
End Sub

Private Sub Comando234_Click()
DoCmd.OpenForm "OrdenDeCompraPlantilla"
[Forms]![OrdenDeCompraPlantilla]![ordenbusqueda] = Me.buscarorden
[Forms]![OrdenDeCompraPlantilla]![afechaorden] = DLookup("[fechaorden]", "[OrdenDeCompra]", "[codordendecompra] = " & Me.buscarorden)
[Forms]![OrdenDeCompraPlantilla]![aproovedor] = DLookup("[codproovedor]", "[OrdenDeCompra]", "[codordendecompra] = " & Me.buscarorden)
[Forms]![OrdenDeCompraPlantilla]![tipopago] = DLookup("[tipopago]", "[OrdenDeCompra]", "[codordendecompra] = " & Me.buscarorden)

If DLookup("[resumendepago]", "[OrdenDeCompra]", "[codordendecompra] = " & Me.buscarorden) <> 0 Then
    [Forms]![OrdenDeCompraPlantilla]![Encabezado_automático0].Caption = "RESUMEN DE PAGO"
    [Forms]![OrdenDeCompraPlantilla]![Comando820].Visible = False
    [Forms]![OrdenDeCompraPlantilla]![Comando106].Visible = False
    [Forms]![OrdenDeCompraPlantilla]![costounitario].Enabled = False
    [Forms]![OrdenDeCompraPlantilla]![cantidadorden].Enabled = False
    [Forms]![OrdenDeCompraPlantilla]![aplicaiva].Enabled = False
End If

[Forms]![OrdenDeCompraPlantilla].Requery
Call Forms("[OrdenDeCompraPlantilla]").actualizarmodopagoinfo
End Sub

Private Sub Comando259_Click()
Me.OrderBy = "[fechaorden] Desc"

End Sub

Private Sub Comando267_Click()
Me.OrderBy = "[codordendecompra] Desc"

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.OrderBy = "[codordendecompra] Desc"

End Sub


