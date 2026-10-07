' ==========================================================
' Modulo  : Gestion Compras
' Tipo    : 1  |  Lineas: 247
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database

Function RequerimientoSemanal(ingrex As String, sem As Integer) As Double

'consumo necesario para la semana vigente aproximadamente, es decir con historial de la semana anterior hacia atras
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'consumo de ingrediente en unas 3 semanas divididos entre 3, considerando de semana anterior hacia atras
miSQL = "SELECT Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])/3) AS PEDIDO," & _
" SubReceta.CodIngrediente" & _
" FROM (DBBatidos" & _
" INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" ON DBBatidos.CodBatido = SubPedido.CodBatido)" & _
" INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha])<" & sem & ", numerosemana([NotaDePedido]![Fecha])>" & sem & "-4," & _
" SubReceta.CodIngrediente" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha])<" & sem & ")<>0) AND ((numerosemana([NotaDePedido]![Fecha])>" & sem & "-4)<>0)" & _
" AND ((SubReceta.CodIngrediente)='" & ingrex & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
RequerimientoSemanal = rst("PEDIDO") * (1 + StockDeSeguridad(ingrex))
rst.Close

Exit Function

Nulo:
RequerimientoSemanal = 0

End Function

Function CondReqSemana(ING As String, Dia As Integer, frec As Integer, sem As Integer) As Double
' Dia es el dia que se va a realizar el ingreso, Lunes = 1, Miercoles = 3, Viernes =2
' Reuqerimiento hasta el dia en mencion, si es lunes a una compra es el 100%, si es miercoles es el 33% si es lunes  dos compras 60%
On Error GoTo Nulo

If Dia = 1 Then  ' lUNES
    If frec = 2 Then '60% ingresa lUNES cuando compra es 2 veces a la semana acorde a ventas
        CondReqSemana = RequerimientoSemanal(ING, sem) * 0.6
    Else
        CondReqSemana = RequerimientoSemanal(ING, sem) / frec
    End If
ElseIf Dia = 3 Then  'mIERCOLES
    CondReqSemana = RequerimientoSemanal(ING, sem) / frec * 2
Else  ' VIERNES
    CondReqSemana = RequerimientoSemanal(ING, sem)
End If

Exit Function

Nulo:
CondReqSemana = 0

End Function

Function CantidadporCompra(frec As Integer, Dia As Integer, ING As String, sem As Integer) As Double
'Cantidad faltante para comprar en el dia de la semana que se hara la compra(Lunes,Miercoles, Viernes)
On Error GoTo Nulo

Dim stock As Double
Dim Cond, Requiere As Double
stock = StockIngrediente(ING, sem) + StockCotizacion(ING, sem)
If frec < Dia Then
    CantidadporCompra = 0
Else
    If CondReqSemana(ING, Dia, frec, sem) > stock Then
        If Dia = 1 Then 'lunes
        CantidadporCompra = CondReqSemana(ING, Dia, frec, sem) - stock
        ElseIf Dia = 3 Then ' miercoles
        CantidadporCompra = CondReqSemana(ING, Dia, frec, sem) - stock - CantidadporCompra(frec, 1, ING, sem)
        Else ' viernes
        CantidadporCompra = CondReqSemana(ING, Dia, frec, sem) - stock - CantidadporCompra(frec, 1, ING, sem) - CantidadporCompra(frec, 3, ING, sem)
        End If
    Else
        CantidadporCompra = 0
    End If
    
End If

Exit Function

Nulo:
CantidadporCompra = 0

End Function
Function CondicionalRequiereCompra(frec As Integer, Dia As Integer, ING As String, sem As Integer) As Long
'Condicional si se requiere compra cierto dia de cierto ingrediente 1: si, 0 : no

Dim stock As Double

stock = StockIngrediente(ING, sem) + StockCotizacion(ING, sem)
If frec < Dia Then
    CondicionalRequiereCompra = 0
Else
    If CondReqSemana(ING, Dia, frec, sem) > stock Then
        CondicionalRequiereCompra = 1
    Else
        CondicionalRequiereCompra = 0
    End If
    
End If

End Function
Function StockDeSeguridad(ING As String) As Double
'Stock de seguridad de ingeidnete

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calculo de stock de seguridad de ingrediente
miSQL = "SELECT DBIngredientes.CodIngrediente, DBIngredientes.StockSeguridad, DBIngredientes.Diallegada, DBIngredientes.ComprasQuincenal FROM DBIngredientes WHERE (((DBIngredientes.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If rst("ComprasQuincenal") = True Then
    StockDeSeguridad = rst("StockSeguridad") + (rst("Diallegada") - 1) * 0.15
    rst.Close
Else
    StockDeSeguridad = rst("StockSeguridad") + (rst("Diallegada") - 1) * 0.15
    rst.Close
End If

Exit Function

Nulo:
StockDeSeguridad = 0

End Function
Function SumaFactura(Fecha As Date, prove As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma toital de factura acorde a proovedor y fecha
miSQL = "SELECT Compras.CodProveedor, Compras.Fecha, Sum(Compras.CostoTotal) AS SumaDeCostoTotal" & _
" FROM Compras " & _
" GROUP BY Compras.CodProveedor, Compras.Fecha HAVING (((Compras.CodProveedor)=" & prove & ") AND ((Compras.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaFactura = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SumaFactura = 0

End Function

Function SumaFacturaGlobal(Fecha As Date, prove As Integer, loc As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma toital de factura acorde a proovedor y fecha
miSQL = "SELECT Compras.CodProveedor, Compras.Fecha, Sum(Compras.CostoTotal) AS SumaDeCostoTotal" & _
" FROM Compras  IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Compras.CodProveedor, Compras.Fecha, Compras.local HAVING (((Compras.CodProveedor)=" & prove & ") AND ((Compras.Fecha)=#" & Fecha & "#) AND ((Compras.local)=" & loc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaFacturaGlobal = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SumaFacturaGlobal = 0

End Function

Function RequerimientoMaximoPorciones(inge As String, por As Integer, rango As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Total Cantidad porciones rango semanas 5 ultimas semanas
miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS semana, NotaDePedido.Anulado, DBIngredientes.CodIngrediente," & _
" mcm([DBIngredientes]![CodIngrediente],[Subreceta]![Cantidad],[Grupos]![Tipo],0) AS Porcion," & _
" Sum([SubPedido]![Cantidad]*mcm([DBIngredientes]![CodIngrediente],[Subreceta]![Cantidad],[Grupos]![Tipo],1)) AS Total" & _
" FROM ((SubReceta INNER JOIN ((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN DBIngredientes ON SubReceta.CodIngrediente = DBIngredientes.CodIngrediente)" & _
" INNER JOIN Grupos ON DBBatidos.CodGrupo = Grupos.CodGrupo" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), NotaDePedido.Anulado, DBIngredientes.CodIngrediente," & _
" mcm([DBIngredientes]![CodIngrediente],[Subreceta]![Cantidad],[Grupos]![Tipo],0)" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha])) Between " & numerosemana(Date) - rango & " And " & numerosemana(Date) - 1 & ")" & _
" AND ((NotaDePedido.Anulado)=0) AND ((DBIngredientes.CodIngrediente)='" & inge & "')" & _
" AND ((mcm([DBIngredientes]![CodIngrediente],[Subreceta]![Cantidad],[Grupos]![Tipo],0))=" & por & "))" & _
" ORDER BY Sum([SubPedido]![Cantidad]*mcm([DBIngredientes]![CodIngrediente],[Subreceta]![Cantidad],[Grupos]![Tipo],1)) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
RequerimientoMaximoPorciones = rst("Total")
rst.Close

Exit Function

Nulo:
RequerimientoMaximoPorciones = 0

End Function

Function MontoOrdenDeCompra(codix As Long) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubOrdenDeCompra.codordendecompra," & _
" Sum([SubOrdenDeCompra]![costounitario]*[SubOrdenDeCompra]![cantidadorden]*IIf([SubOrdenDeCompra]![aplicaiva]=0,1,1.15)) AS Expr1" & _
" FROM SubOrdenDeCompra" & _
" GROUP BY SubOrdenDeCompra.codordendecompra" & _
" HAVING (((SubOrdenDeCompra.codordendecompra)=" & codix & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MontoOrdenDeCompra = rst("Expr1")
rst.Close

Exit Function

Nulo:
MontoOrdenDeCompra = 0

End Function


Function statuscomprasucursal(codi As Long, sucu As Integer) As Integer
'0 no existe orden de compra por ende no esta egnerado, 1 ya se hiso pago
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT AprobacionPagoComprasSucursales.CodIngresoAlmacen, AprobacionPagoComprasSucursales.Sucursal," & _
" AprobacionPagoComprasSucursales.codordendecompra" & _
" FROM AprobacionPagoComprasSucursales" & _
" WHERE (((AprobacionPagoComprasSucursales.CodIngresoAlmacen)=" & codi & ") AND ((AprobacionPagoComprasSucursales.Sucursal)=" & sucu & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
statuscomprasucursal = IIf(rst("codordendecompra") >= 0, 1, 0)
rst.Close

Exit Function

Nulo:
statuscomprasucursal = 0

End Function


