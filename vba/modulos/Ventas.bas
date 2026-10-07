' ==========================================================
' Modulo  : Ventas
' Tipo    : 1  |  Lineas: 382
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Function VentasXTipoPorcionDia(diax As Date, porci As Integer) As Double
'Ventas de alguna procion en el dia

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las ventas totales
miSQL = "SELECT NotaDePedido.Anulado, NotaDePedido.Fecha, SubReceta.codporcion," & _
" Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]/DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])) AS Total" & _
" FROM ((SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, NotaDePedido.Fecha, SubReceta.codporcion" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((NotaDePedido.Fecha)=#" & diax & "#) AND ((SubReceta.codporcion)=" & porci & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasXTipoPorcionDia = rst("Total")
rst.Close

Exit Function

Nulo:
VentasXTipoPorcionDia = 0

End Function

Function VentasXTipoPorcionSemana(semax As Integer, porci As Integer) As Double
'Ventas de alguna procion en el dia

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las ventas totales
miSQL = "SELECT NotaDePedido.Anulado, numerosemana([NotaDePedido]![Fecha]) AS semana, SubReceta.codporcion," & _
" Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]/DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])) AS Total" & _
" FROM ((SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, numerosemana([NotaDePedido]![Fecha]), SubReceta.codporcion" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((numerosemana([NotaDePedido]![Fecha]))=" & semax & ") AND ((SubReceta.codporcion)=" & porci & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasXTipoPorcionSemana = rst("Total")
rst.Close

Exit Function

Nulo:
VentasXTipoPorcionSemana = 0

End Function

Function VentasXTipoNoPorcionDia(diax As Date, noporci As String) As Double
'Ventas de alguna no procion en el dia
'noporci es codigo de ingrediente

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las ventas totales
miSQL = "SELECT NotaDePedido.Anulado, IsNull([SubReceta]![codporcion]) AS noesporcion, NotaDePedido.Fecha, SubReceta.CodIngrediente," & _
" Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]) AS Total" & _
" FROM ((SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, IsNull([SubReceta]![codporcion]), NotaDePedido.Fecha, SubReceta.CodIngrediente" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((IsNull([SubReceta]![codporcion]))<>0)" & _
" AND ((NotaDePedido.Fecha)=#" & diax & "#) AND ((SubReceta.CodIngrediente)='" & noporci & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasXTipoNoPorcionDia = rst("Total")
rst.Close

Exit Function

Nulo:
VentasXTipoNoPorcionDia = 0

End Function


Function VentasXTipoNoPorcionSemana(semax As Integer, noporci As String) As Double
'Ventas de alguna no procion en el dia
'noporci es codigo de ingrediente
'filtrado no lee consumo de porciones aunque sea ingrediente igual

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las ventas totales
miSQL = "SELECT NotaDePedido.Anulado, IsNull([SubReceta]![codporcion]) AS noesporcion," & _
" numerosemana([NotaDePedido]![Fecha]) AS semana, SubReceta.CodIngrediente," & _
" Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]) AS Total" & _
" FROM ((SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, IsNull([SubReceta]![codporcion]), numerosemana([NotaDePedido]![Fecha]), SubReceta.CodIngrediente" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((IsNull([SubReceta]![codporcion]))<>0)" & _
" AND ((numerosemana([NotaDePedido]![Fecha]))=" & semax & ") AND ((SubReceta.CodIngrediente)='" & noporci & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasXTipoNoPorcionSemana = rst("Total")
rst.Close

Exit Function

Nulo:
VentasXTipoNoPorcionSemana = 0

End Function

Function VentasXTipoNoPorcionMaximoSemana(semax As Integer, noporci As String, rang As Integer) As Double
'Ventas de alguna no procion el maximo en las rang ultimas semanas invcluyendo la semana semax
'noporci es codigo de ingrediente
'filtrado no lee consumo de porciones aunque sea ingrediente igual

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las ventas totales
miSQL = "SELECT NotaDePedido.Anulado, IsNull([SubReceta]![codporcion]) AS noesporcion," & _
" numerosemana([NotaDePedido]![Fecha]) AS semana, SubReceta.CodIngrediente, Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]) AS Total" & _
" FROM ((SubPedido INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, IsNull([SubReceta]![codporcion]), numerosemana([NotaDePedido]![Fecha]), SubReceta.CodIngrediente" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((IsNull([SubReceta]![codporcion]))<>0)" & _
" AND ((numerosemana([NotaDePedido]![Fecha])) Between " & semax - rang + 1 & " And " & semax & ") AND ((SubReceta.CodIngrediente)='" & noporci & "'))" & _
" ORDER BY Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasXTipoNoPorcionMaximoSemana = rst("Total")
rst.Close

Exit Function

Nulo:
VentasXTipoNoPorcionMaximoSemana = 0

End Function


Function cantidadpedidospendientescentral(sucu As Integer, fechac As Date) As Integer
'cantidad pedidos generados en la central aprobadaos y aun noa gendidos

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusPedidosCentralDeliveryCentral.Fecha, StatusPedidosCentralDeliveryCentral.Sucursal," & _
" Sum(IIf(codigopedidosucursaldepedidocentral([StatusPedidosCentralDeliveryCentral]![CodPedidoCentral])=0,1,0)) AS total" & _
" FROM StatusPedidosCentralDeliveryCentral" & _
" GROUP BY StatusPedidosCentralDeliveryCentral.Fecha, StatusPedidosCentralDeliveryCentral.Sucursal" & _
" HAVING (((StatusPedidosCentralDeliveryCentral.Fecha)=#" & fechac & "#) AND ((StatusPedidosCentralDeliveryCentral.Sucursal)=" & sucu & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cantidadpedidospendientescentral = rst("total")
rst.Close

Exit Function

Nulo:
cantidadpedidospendientescentral = 0

End Function

Function codigopedidosucursaldepedidocentral(pedicentral As Long) As Long
'busca el peduido generado en la sucursal de un pedido de centrl

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusPedidosCentral.CodPedidoCentral, StatusPedidosCentral.CodPedidoSucursal" & _
" FROM StatusPedidosCentral" & _
" WHERE (((StatusPedidosCentral.CodPedidoCentral)=" & pedicentral & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
codigopedidosucursaldepedidocentral = rst("CodPedidoSucursal")
rst.Close

Exit Function

Nulo:
codigopedidosucursaldepedidocentral = 0

End Function

Function statusaprobaciondeliverycentral(pedicentral As Long) As Boolean
'si esta aprobado o no desde la central

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusPedidosCentral.CodPedidoCentral, StatusPedidosCentral.AprobadoCentral FROM StatusPedidosCentral" & _
" WHERE (((StatusPedidosCentral.CodPedidoCentral)=" & pedicentral & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
statusaprobaciondeliverycentral = rst("AprobadoCentral")
rst.Close

Exit Function

Nulo:
statusaprobaciondeliverycentral = 0

End Function

Sub moverfacturacentralasucursal(pedicentral As Long, pedisucursal As Long, condi As Integer)
'confi :0 sin motorizado, 1: con motorizado

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT NotaDePedidoDeliveryCentral.Modalidad, NotaDePedidoDeliveryCentral.CodCliente, NotaDePedidoDeliveryCentral.[Nombre Provicional]," & _
" NotaDePedidoDeliveryCentral.POS, NotaDePedidoDeliveryCentral.Delivery, NotaDePedidoDeliveryCentral.CodMotorizado," & _
" NotaDePedidoDeliveryCentral.Transferencia, NotaDePedidoDeliveryCentral.DeliveryComision," & _
" NotaDePedidoDeliveryCentral.Observaciones, NotaDePedidoDeliveryCentral.recibecor, NotaDePedidoDeliveryCentral.recibedol" & _
" FROM NotaDePedidoDeliveryCentral WHERE (((NotaDePedidoDeliveryCentral.CodPedido)=" & pedicentral & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If condi = 0 Then 'sin mototizado
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO NotaDePedido (CodPedido, Fecha, Hora, Modalidad, CodCliente, [Nombre Provicional]," & _
    " senuelo, POS, Delivery," & _
    " Transferencia, DeliveryComision," & _
    " HoraCreado, HoraIngresoProducto, Observaciones, recibecor, recibedol)" & _
    " VALUES (" & pedisucursal & ", #" & Date & "#, #" & Time & "#, " & IIf(rst("Modalidad"), "True", "False") & ", " & rst("CodCliente") & ", '" & rst("[Nombre Provicional]") & "'," & _
    " 0, " & IIf(rst("POS"), "True", "False") & ", " & rst("Delivery") & "," & _
    " " & IIf(rst("Transferencia"), "True", "False") & ", " & IIf(rst("DeliveryComision"), "True", "False") & "," & _
    " #" & Time & "#, #" & Time & "#, '" & rst("Observaciones") & "', " & rst("recibecor") & ", " & rst("recibedol") & ")"
    DoCmd.SetWarnings True
Else
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO NotaDePedido (CodPedido, Fecha, Hora, Modalidad, CodCliente, [Nombre Provicional]," & _
    " senuelo, POS, Delivery, CodMotorizado," & _
    " Transferencia, DeliveryComision," & _
    " HoraCreado, HoraIngresoProducto, Observaciones, recibecor, recibedol)" & _
    " VALUES (" & pedisucursal & ", #" & Date & "#, #" & Time & "#, " & IIf(rst("Modalidad"), "True", "False") & ", " & rst("CodCliente") & ", '" & rst("[Nombre Provicional]") & "'," & _
    " 0, " & IIf(rst("POS"), "True", "False") & ", " & rst("Delivery") & ", " & rst("CodMotorizado") & "," & _
    " " & IIf(rst("Transferencia"), "True", "False") & ", " & IIf(rst("DeliveryComision"), "True", "False") & "," & _
    " #" & Time & "#, #" & Time & "#, '" & rst("Observaciones") & "', " & rst("recibecor") & ", " & rst("recibedol") & ")"
    DoCmd.SetWarnings True
End If

rst.Close

Exit Sub

Nulo:
MsgBox "No existe pedido"

End Sub

Sub moverdetallefacturacentralasucursal(pedicentral As Long, pedisucursal As Long)


On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contax As Integer
Dim contpe As Integer
Dim colax As Long
Dim vincux As Long

miSQL = "SELECT IIf(IsNull([SubPedidoDeliveryCentral]![Vinculo]),[SubPedidoDeliveryCentral]![CodSubPedido],IIf([SubPedidoDeliveryCentral]![Vinculo]=0,[SubPedidoDeliveryCentral]![CodSubPedido],[SubPedidoDeliveryCentral]![Vinculo])) AS orde," & _
" SubPedidoDeliveryCentral.CodSubPedido, SubPedidoDeliveryCentral.CodPedido," & _
" SubPedidoDeliveryCentral.CodBatido, SubPedidoDeliveryCentral.Cantidad," & _
" SubPedidoDeliveryCentral.CodPromocion, SubPedidoDeliveryCentral.Observaciones, SubPedidoDeliveryCentral.DetallesAdicionales, SubPedidoDeliveryCentral.SinAzucar," & _
" SubPedidoDeliveryCentral.Azucar, SubPedidoDeliveryCentral.Vinculo, SubPedidoDeliveryCentral.Empaque" & _
" FROM SubPedidoDeliveryCentral" & _
" WHERE (((SubPedidoDeliveryCentral.CodPedido) = " & pedicentral & "))" & _
" ORDER BY IIf(IsNull([SubPedidoDeliveryCentral]![Vinculo]),[SubPedidoDeliveryCentral]![CodSubPedido],IIf([SubPedidoDeliveryCentral]![Vinculo]=0,[SubPedidoDeliveryCentral]![CodSubPedido],[SubPedidoDeliveryCentral]![Vinculo]))," & _
" SubPedidoDeliveryCentral.CodSubPedido"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
contax = rst.RecordCount
rst.MoveFirst

For contpe = 1 To contax
    
    If rst("CodBatido") <> "C167" Then ' si es delivery departamental no lo copia
        
        If rst("orde") = rst("CodSubPedido") Then
            colax = ultimosubpedidofacturado() + 1
            vincux = 0
        Else
            vincux = colax
        End If
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido," & _
        " CodPromocion, Observaciones, DetallesAdicionales, SinAzucar," & _
        " Azucar, Vinculo, Empaque)" & _
        " VALUES ('" & rst("CodBatido") & "', " & rst("Cantidad") & ", " & pedisucursal & "," & _
        " " & rst("CodPromocion") & ", '" & rst("Observaciones") & "', '" & rst("DetallesAdicionales") & "', " & IIf(rst("SinAzucar"), "True", "False") & "," & _
        " '" & rst("Azucar") & "', " & vincux & ", " & IIf(rst("Empaque"), "True", "False") & ")"
        DoCmd.SetWarnings True
    
    End If
    
    rst.MoveNext
Next contpe

rst.Close

Exit Sub

Nulo:
MsgBox "No existe pedido"

End Sub

Function ventastotalesservicioatencionalcliente(fdesde As Date, fhasta As Date) As Double
'total de ventas gesitonadas en un lapso de tiempo proa tencon al cliente

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT StatusPedidosCentral.AprobadoCentral, NotaDePedido.Anulado," & _
" [NotaDePedido]![Fecha]>=#" & fdesde & "# And [NotaDePedido]![Fecha]<=#" & fhasta & "# AS intervalo," & _
" Sum([NotaDePedido]![TotalGuardado]) AS Total" & _
" FROM ((StatusPedidosCentral INNER JOIN NotaDePedido" & _
" ON StatusPedidosCentral.CodPedidoCentral = NotaDePedido.CodPedido)" & _
" INNER JOIN ClientesDelivery ON NotaDePedido.CodClientesDelivery = ClientesDelivery.CodClientesDelivery)" & _
" GROUP BY StatusPedidosCentral.AprobadoCentral, NotaDePedido.Anulado," & _
" [NotaDePedido]![Fecha]>=#" & fdesde & "# And [NotaDePedido]![Fecha]<=#" & fhasta & "#" & _
" HAVING (((StatusPedidosCentral.AprobadoCentral)<>0) AND ((NotaDePedido.Anulado)=0)" & _
" AND (([NotaDePedido]![Fecha]>=#" & fdesde & "# And [NotaDePedido]![Fecha]<=#" & fhasta & "#)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ventastotalesservicioatencionalcliente = rst("Total")
rst.Close

Exit Function

Nulo:
ventastotalesservicioatencionalcliente = 0

End Function

Function VentasTotalesGuardadoSemana(semanacal As Integer) As Double

'On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la venta acumulada del dia acorde a semana especifica

miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Semana, Sum(Round([TotalGuardado],0)) AS SumaDeTotalGuardado" & _
" FROM NotaDePedido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha])" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & semanacal & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VentasTotalesGuardadoSemana = rst("SumaDeTotalGuardado")
rst.Close
     
Exit Function

Nulo:
VentasTotalesGuardadoSemana = 0

End Function

Function montocuponvigente(cupon As String) As Integer

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer

miSQL = "SELECT cupones_sucursales.numero_cupon, cupones_sucursales.monto," & _
" cupones_sucursales.aplicado, cupones_sucursales.fecha_caducidad" & _
" FROM cupones_sucursales" & _
" WHERE (((cupones_sucursales.numero_cupon)='" & cupon & "')" & _
" AND ((cupones_sucursales.aplicado)=0)" & _
" AND ((cupones_sucursales.fecha_caducidad)>=Date()))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
montocuponvigente = rst("monto")
rst.Close

Exit Function

Nulo:
montocuponvigente = 0
End Function
