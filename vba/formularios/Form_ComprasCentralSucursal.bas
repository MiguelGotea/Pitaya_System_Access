' ==========================================================
' Modulo  : Form_ComprasCentralSucursal
' Tipo    : 100  |  Lineas: 209
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:19
' ==========================================================

Option Compare Database
Function proovedoresseleccionadosgeneracionresumendepago() As Integer
'si no hay mas de 2 proovedores seleccionados manda el nunero de provedor, si hay mas de dos manda 0
On Error Resume Next
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT ComprasSucursal.Elegir, ComprasSucursal.CodProveedor" & _
" FROM ComprasSucursal" & _
" GROUP BY ComprasSucursal.Elegir, ComprasSucursal.CodProveedor" & _
" HAVING (((ComprasSucursal.Elegir)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
If rst.RecordCount > 1 Then
    proovedoresseleccionadosgeneracionresumendepago = 0
Else
    proovedoresseleccionadosgeneracionresumendepago = rst("CodProveedor")
End If

rst.Close
Exit Function

Nulo:
proovedoresseleccionadosgeneracionresumendepago = 0
End Function

Function proovedorelegido() As Integer
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT ComprasSucursal.Tipo, statuscomprasucursal([ComprasSucursal]![CodIngresoAlmacen]," & _
" [ComprasSucursal]![Sucursal]) AS status, ComprasSucursal.Elegir, ComprasSucursal.CodProveedor" & _
" FROM ComprasSucursal" & _
" GROUP BY ComprasSucursal.Tipo, statuscomprasucursal([ComprasSucursal]![CodIngresoAlmacen],[ComprasSucursal]![Sucursal])," & _
" ComprasSucursal.Elegir, ComprasSucursal.CodProveedor" & _
" HAVING (((ComprasSucursal.Tipo)='CENTRAL')" & _
" AND ((statuscomprasucursal([ComprasSucursal]![CodIngresoAlmacen],[ComprasSucursal]![Sucursal]))=0)" & _
" AND ((ComprasSucursal.Elegir)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
proovedorelegido = rst("CodProveedor")
rst.Close
Exit Function

Nulo:
proovedorelegido = 0
End Function

Sub llenarocproductosseleccionados(Orden As Long, sucux As Integer)
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cante As Integer

Dim fcodigo As Long
Dim fprodu As Integer
Dim fcanti As Double
Dim fcosto As Double
Dim Cal As Double

miSQL = "SELECT ComprasSucursal.CodIngresoAlmacen, ComprasSucursal.Fecha," & _
" ComprasSucursal.CodProveedor, ComprasSucursal.CodCotizacion, ComprasSucursal.Cantidad," & _
" ComprasSucursal.CostoTotal, ComprasSucursal.Tipo, ComprasSucursal.Elegir" & _
" FROM ComprasSucursal" & _
" WHERE ((ComprasSucursal.Elegir) <>0)" & _
" ORDER BY ComprasSucursal.CodIngresoAlmacen, ComprasSucursal.Fecha DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cante = rst.RecordCount
rst.MoveFirst

For I = 1 To cante
    If I >= Me.icoti Then
        fcodigo = rst("CodIngresoAlmacen")
        fprodu = rst("CodCotizacion")
        fcanti = rst("Cantidad")
        fcosto = rst("CostoTotal") / fcanti
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO SubOrdenDeCompra(codordendecompra, codcotizacion, cantidadorden, destino, costounitario)" & _
        " values (" & Orden & ", " & fprodu & ", " & fcanti & ", " & sucux & ", " & fcosto & ")"
        DoCmd.RunSQL "INSERT INTO AprobacionPagoComprasSucursales(CodIngresoAlmacen, Sucursal, Aprobado, codordendecompra)" & _
        " values (" & fcodigo & ", " & sucux & ",# " & Now & "#, " & Orden & ")"
        DoCmd.SetWarnings True
        
    End If
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call llenarocproductosseleccionados(Orden, sucux)
End Sub


Private Sub Comando155_Click()
On Error Resume Next
Call importartablaespecifica("Pitaya" & Me.asucursal & "_DB", "Compras", "ComprasSucursal", 1)

DoCmd.SetWarnings False
DoCmd.RunSQL "ALTER TABLE ComprasSucursal ADD COLUMN Sucursal Integer, Elegir YesNo"
'DoCmd.RunSQL "ALTER TABLE ComprasSucursal ADD COLUMN Elegir YesNo"
DoCmd.RunSQL "UPDATE ComprasSucursal SET ComprasSucursal.Sucursal = " & Me.asucursal
DoCmd.SetWarnings True

Me.Requery
End Sub

Private Sub Comando939_Click()

Dim prove As Integer
prove = proovedoresseleccionadosgeneracionresumendepago()
Me.Requery

If IsNull(Me.asucursal) Then
    MsgBox "Seleccionar sucursal para elegir facturas"
    Exit Sub
End If

If MsgBox("Desea generar Resumen de pago de los productos seleccionados? una vez generado el resumen de pago ya no se podra editar las facturas registradas", vbYesNo, "CONFIRMACION") <> vbYes Then
    Exit Sub
End If

If prove = 0 Then
    MsgBox "No se pueden seleccionar productos de diferenctes proovedores y por lo menos seleccionar un producto para generar resumen de pago"
    Exit Sub
End If

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO OrdenDeCompra(codproovedor, fechaorden, resumendepago)" & _
" values (" & prove & ", #" & Date & "#, -1)"
DoCmd.SetWarnings True

Dim codordencompra As Long
codordencompra = ultimaordendecompra()
Call llenarocproductosseleccionados(codordencompra, Me.asucursal)

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE ComprasSucursal SET ComprasSucursal.Elegir = 0"
DoCmd.SetWarnings True

Me.Requery

DoCmd.OpenForm "OrdenDeCompraPlantilla"

[Forms]![OrdenDeCompraPlantilla]![ordenbusqueda] = codordencompra
[Forms]![OrdenDeCompraPlantilla]![afechaorden] = Date
[Forms]![OrdenDeCompraPlantilla]![aproovedor] = prove


[Forms]![OrdenDeCompraPlantilla]![Encabezado_automático0].Caption = "RESUMEN DE PAGO"
[Forms]![OrdenDeCompraPlantilla]![Comando820].Visible = False
[Forms]![OrdenDeCompraPlantilla]![Comando106].Visible = False
[Forms]![OrdenDeCompraPlantilla]![costounitario].Enabled = False
[Forms]![OrdenDeCompraPlantilla]![cantidadorden].Enabled = False
[Forms]![OrdenDeCompraPlantilla]![aplicaiva].Enabled = False

[Forms]![OrdenDeCompraPlantilla].Requery

End Sub

Private Sub Form_Close()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM ComprasSucursal"
DoCmd.SetWarnings True
End Sub

Private Sub Form_Open(Cancel As Integer)


Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
End Sub




Private Sub Modalidad_Click()
'IIf(Me.Modalidad = "PENDIENTE", -1, 0)
Dim cambio As Integer

If Me.Ingresado = 0 Then
    MsgBox "no se puede generar resumen de pago de facturas que no hayan sido ingresadas por cierre en la sucursal"
    Exit Sub
End If
If IsNull(Me.asucursal) Then
    MsgBox "Seleccionar sucursal para elegir facturas"
    Exit Sub
End If

If Me.CodProveedor = proovedorelegido() Or proovedorelegido() = 0 Then 'se elige mismo proovedor o no hay aun proovedor elegido

    cambio = IIf(Me.Modalidad = "PAGAR", 0, -1)
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "UPDATE ComprasSucursal SET ComprasSucursal.Elegir = " & cambio & _
    " WHERE ComprasSucursal.NumeroFactura = '" & Me.NumeroFactura & "'"
    DoCmd.SetWarnings True
    Me.Requery

Else
    MsgBox "Ya existe un proovedor seleccionado para hacer Resumen de Pago, quitar selecicon del otro proovedor para elegir las facturas de uno nuevo"
    Exit Sub
End If
End Sub
