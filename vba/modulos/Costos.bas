' ==========================================================
' Modulo  : Costos
' Tipo    : 1
' Lineas  : 626
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:22
' ==========================================================
Option Compare Database
' NUEVAS FORMULAS

Function CostoUnitarioCotizacionxCompra(coti As Integer, sem As Integer) As Double
'Costo Unitario promedio de una cotizacion de la semana  segun las compras de esa semana
'SI HAY REGISTRO DE COMPRA LOCAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([Compras]![Fecha]) AS sema, Compras.CodCotizacion," & _
" Sum(Compras.Cantidad) AS SumaDeCantidad, Sum(Compras.CostoTotal) AS SumaDeCostoTotal" & _
" FROM compras" & _
" GROUP BY numerosemana([Compras]![Fecha]), Compras.CodCotizacion" & _
" HAVING (((numerosemana([Compras]![Fecha]))<=" & sem & ") AND ((Compras.CodCotizacion)=" & coti & "))" & _
" ORDER BY numerosemana([Compras]![Fecha]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoUnitarioCotizacionxCompra = rst("SumaDeCostoTotal") / rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
CostoUnitarioCotizacionxCompra = 0

End Function

Function CostoUnitarioIngredientexCompra(insu As String, sem As Integer) As Double
'Costo Unitario promedio de una ingrediente de la semana  segun las compras de esa semana
'SI HAY REGISTRO DE INGRESO LOCAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([IngresosPitaya]![Fecha]) AS sema," & _
" Sum([IngresosPitaya]![Cantidad]*conversioncalculado([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS Ingreso," & _
" Sum([IngresosPitaya]![Cantidad]*CostoUnitarioCotizacionxCompra([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS Valor" & _
" FROM (IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion)" & _
" INNER JOIN DBIngredientes ON Cotizaciones.CodIngrediente = DBIngredientes.CodIngrediente" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([IngresosPitaya]![Fecha])" & _
" HAVING (((Cotizaciones.CodIngrediente) = '" & insu & "') And ((numerosemana([IngresosPitaya]![Fecha])) <= " & sem & "))" & _
" ORDER BY numerosemana([IngresosPitaya]![Fecha]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoUnitarioIngredientexCompra = rst("Valor") / rst("Ingreso")
rst.Close

Exit Function

Nulo:
CostoUnitarioIngredientexCompra = 0

End Function
' FIND E NUEVAS FORMULAS

Function CostoUnitarioUnidad(coti As Integer, sem As Integer) As Double
'Costo Unitario promedio de una cotizacion de la semana mencionada hacia abajo

On Error GoTo Nulo
Dim Valor As Double

If esporcionablecotizacion(coti) <> 0 Then 'CUando es una porcion
    CostoUnitarioUnidad = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti) * FRCostoIngredienteGlobal(DLookup("[CodIngrediente]", "Cotizaciones", "[CodCotizacion] = " & coti), sem)
    Exit Function
End If

If codigoLocal() = 0 Then 'Sistema local
    CostoUnitarioUnidad = CostoUnitarioUnidadLocal(coti, sem)
    Exit Function
End If

If DLookup("[Global]", "Cotizaciones", "[CodCotizacion] = " & coti) = 0 Then 'global = 0 'Productos que se compran para local solamente
    Valor = CostoUnitarioUnidadLocal(coti, sem)
    If Valor = 0 Then
        Valor = CostoUnitarioUnidadMixed(coti, sem)
    End If
Else    ' COsto global, que se compra para todos los locales
    Valor = CostoUnitarioUnidadMixed(coti, sem)
End If

CostoUnitarioUnidad = Valor

Exit Function

Nulo:
'MsgBox "error al hallar valor de cotizacion"
CostoUnitarioUnidad = 0
End Function

Function CostoUnitarioUnidadLocal(coti As Integer, sem As Integer) As Double
'Costo Unitario promedio de una cotizacion de la semana mencionada hacia abajo EN LOCAL
'SI HAY REGISTRO DE COMPRA LOCAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

' "Costo LOCAL sem" & sem & "coti " & coti
'Calcula el coso unitario de cotizacion promedio en local
miSQL = "SELECT TOP 1 numerosemana([Compras]![Fecha]) AS semana, Compras.CodCotizacion," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Sum(Compras.Cantidad) AS SumaDeCantidad" & _
" FROM Compras" & _
" GROUP BY numerosemana([Compras]![Fecha]), Compras.CodCotizacion" & _
" HAVING (((numerosemana([Compras]![Fecha])) <= " & sem & ") And ((Compras.CodCotizacion) = " & coti & ") And" & _
" ((Sum(Compras.CostoTotal)) <> 0) And ((Sum(Compras.Cantidad)) <> 0))" & _
" ORDER BY numerosemana([Compras]![Fecha]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoUnitarioUnidadLocal = rst("SumaDeCostoTotal") / rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
CostoUnitarioUnidadLocal = 0

End Function

Function CostoUnitarioUnidadMixed(coti As Integer, sem As Integer) As Double
'Costo Unitario promedio de una cotizacion de la semana mencionada hacia abajo EN MIXED
'CUANDO NON HAY VALOR EN BASE DE DATOS LOCAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

' "Costo mix sem" & sem & "coti " & coti
'Calcula el coso unitario de cotizacion promedio en mixed

miSQL = "SELECT TOP 1 numerosemana([Compras]![Fecha]) AS semana, Compras.CodCotizacion," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Sum(Compras.Cantidad) AS SumaDeCantidad, IIf([Compras]![local]=codigolocal(),1,0) AS cond" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY numerosemana([Compras]![Fecha]), Compras.CodCotizacion, IIf([Compras]![local]=codigolocal(),1,0)" & _
" HAVING (((numerosemana([Compras]![Fecha])) <= " & sem & ") And ((Compras.CodCotizacion) = " & coti & ") And" & _
" ((Sum(Compras.CostoTotal)) <> 0) And ((Sum(Compras.Cantidad)) <> 0))" & _
" ORDER BY numerosemana([Compras]![Fecha]) DESC , IIf([Compras]![local]=codigolocal(),1,0) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoUnitarioUnidadMixed = rst("SumaDeCostoTotal") / rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
CostoUnitarioUnidadMixed = 0

End Function

Function CostoIngredienteProcesamiento(ING As String, sem As Integer, modo As Integer) As Double
'Costo de procesamiento de un ingredinte en una semana especifica
'COSO DE LA CANTIDAD DE INGREDIENTE PROCESADO EN LA SEMANA Y LA CANTIDAD PROCESADA EN LA SEMANA4
'UNIDADES CON CONVERSION = 0   ' INGRESO POR PROCESAMIENTO

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular los valores de total procesdo de ingrediente (1) y costo total de lo procesado en cotizacion (2)
miSQL = "SELECT numerosemana([Procesamiento]![Fecha]) AS semana, Cotizaciones.CodIngrediente, Sum(Procesamiento.MedidaFinal) AS SumaDeMedidaFinal," & _
" Sum([Procesamiento]![Cantidad]*CostoUnitarioUnidad([Procesamiento]![CodCotizacion],numerosemana([Procesamiento]![Fecha]))) AS CostoTotal" & _
" FROM Procesamiento INNER JOIN Cotizaciones ON Procesamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([Procesamiento]![Fecha]), Cotizaciones.CodIngrediente" & _
" HAVING (((numerosemana([Procesamiento]![Fecha]))=" & sem & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If modo = 1 Then
CostoIngredienteProcesamiento = rst("SumaDeMedidaFinal")
rst.Close
Else
CostoIngredienteProcesamiento = rst("CostoTotal")
rst.Close
End If


Exit Function

Nulo:
CostoIngredienteProcesamiento = 0

End Function

Function CostoIngredienteIngreso(ING As String, sem As Integer, modo As Integer) As Double
'Costo de ingresos de un ingredinte en una semana especifica
'COSTO DE LO Q INGRESA EN LA SEMANA Y TOTAL EN VALOR ING DE LO QUE INGRESA

' UNIDADES CON CONVERSION NO 0  ' INGRESO NATURAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular los valores de total ingresado de ingrediente (1) y costo total de lo ingresado en cotizacion (2)
miSQL = "SELECT numerosemana([IngresosPitaya]![Fecha]) AS semana, Cotizaciones.CodIngrediente, [Cotizaciones]![Conversion]=0 AS conversion," & _
" Sum([IngresosPitaya]![Cantidad]*ConversionCalculado([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS IngresoTotal," & _
" Sum([IngresosPitaya]![Cantidad]*CostoUnitarioUnidad([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS CostoTotal" & _
" FROM IngresosPitaya" & _
" INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([IngresosPitaya]![Fecha]), Cotizaciones.CodIngrediente, [Cotizaciones]![Conversion]=0" & _
" HAVING (((numerosemana([IngresosPitaya]![Fecha]))=" & sem & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "') AND (([Cotizaciones]![Conversion]=0)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If modo = 1 Then
CostoIngredienteIngreso = rst("IngresoTotal")
rst.Close
Else
CostoIngredienteIngreso = rst("CostoTotal")
rst.Close
End If


Exit Function

Nulo:
CostoIngredienteIngreso = 0

End Function

Function CostoIngredienteGlobal(ING As String, sem As Integer) As Double
' Costo Unitario de ingrediente acorde a ingresos y procesamientos de una semana especifica
' SI NO HAY DATOS EN LA SEMANA AGARRA DE SEMANA ATERIOR
'Aplica abajo de la semaana 160
On Error GoTo Nulo

If sem = 0 Then
    CostoIngredienteGlobal = 0
    Exit Function
End If

If sem >= 160 Then
    If StockCotizacionPorcionable(ING, sem) + StockIngrediente(ING, sem) = 0 Then
        CostoIngredienteGlobal = (CostoIngredienteProcesamiento(ING, sem, 2) + CostoIngredienteIngreso(ING, sem, 2)) / (CostoIngredienteProcesamiento(ING, sem, 1) + CostoIngredienteIngreso(ING, sem, 1))
    Else
        CostoIngredienteGlobal = ((StockCotizacionPorcionable(ING, sem) + StockIngrediente(ING, sem)) * FRCostoIngredienteGlobal(ING, sem - 1) + CostoIngredienteProcesamiento(ING, sem, 2) + CostoIngredienteIngreso(ING, sem, 2)) / (CostoIngredienteProcesamiento(ING, sem, 1) + CostoIngredienteIngreso(ING, sem, 1) + (StockCotizacionPorcionable(ING, sem) + StockIngrediente(ING, sem)))
    End If
Else
    CostoIngredienteGlobal = (CostoIngredienteProcesamiento(ING, sem, 2) + CostoIngredienteIngreso(ING, sem, 2)) / (CostoIngredienteProcesamiento(ING, sem, 1) + CostoIngredienteIngreso(ING, sem, 1))
End If

Exit Function

Nulo:
CostoIngredienteGlobal = CostoIngredienteGlobal(ING, sem - 1)

End Function

Function CostoIngredienteGlobalxCompras(inge As String, sem As Integer) As Double
'suma de total ingrediente de la semana ycosto total de la semana tabla de todos sucursales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([Compras]![Fecha]) AS sem," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal," & _
" Sum([Compras]![Cantidad]*ConversionCalculadoGlobal([Compras]![CodCotizacion],numerosemana([Compras]![Fecha]))) AS cantitotal" & _
" FROM Compras INNER JOIN Cotizaciones ON Compras.CodCotizacion = Cotizaciones.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Compras]![Fecha])" & _
" HAVING (((Cotizaciones.CodIngrediente) = '" & inge & "') And ((numerosemana([compras]![Fecha])) < " & sem & "))" & _
" ORDER BY numerosemana([Compras]![Fecha]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
CostoIngredienteGlobalxCompras = rst("SumaDeCostoTotal") / rst("cantitotal")
rst.Close

Exit Function

Nulo:
CostoIngredienteGlobalxCompras = 0

End Function

Function CostoIngredienteGlobalxCompras2(inge As String, sem As Integer) As Double
'suma de total ingrediente de la semana ycosto total de la semana tabla de todos sucursales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim canti As Integer

Dim tota As Double
Dim costot As Double
Dim ttota As Double
Dim tcost As Double
Dim cotiz As Integer
Dim tempo As Double

costot = 0
tota = 0

miSQL = "SELECT Cotizaciones.CodIngrediente, Cotizaciones.CodCotizacion" & _
" FROM Cotizaciones WHERE (((Cotizaciones.CodIngrediente)='" & inge & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst

For I = 1 To canti
    cotiz = rst("CodCotizacion")
    tempo = SumaCantiCostoCompraGlobal(cotiz, sem, tcost, ttota)
    tota = tota + ttota * ConversionCalculadoGlobal(cotiz, sem)
    costot = costot + tcost
    rst.MoveNext
Next I

rst.Close

CostoIngredienteGlobalxCompras2 = costot / tota
Exit Function

Nulo:
CostoIngredienteGlobalxCompras2 = CostoIngredienteGlobalxCompras2(inge, sem - 1)

End Function

Function SumaCantiCostoCompraGlobal(cotit As Integer, semi As Integer, ByRef mont As Double, ByRef acant As Double) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Compras.CodCotizacion, numerosemana([Compras]![Fecha]) AS sema," & _
" Sum(Compras.Cantidad) AS SumaDeCantidad, Sum(Compras.CostoTotal) AS SumaDeCostoTotal" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Compras.CodCotizacion, numerosemana([Compras]![Fecha])" & _
" HAVING (((compras.CodCotizacion) = " & cotit & ") And ((numerosemana([compras]![Fecha])) = " & semi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
mont = rst("SumaDeCostoTotal")
acant = rst("SumaDeCantidad")
rst.Close
SumaCantiCostoCompraGlobal = 0
Exit Function

Nulo:
mont = 0
acant = 0
SumaCantiCostoCompraGlobal = 0

End Function

Function CostoInventarioCotizacion(sem As Integer) As Double
'Costo de los items de cotizacion realizados los domingos, de una semana especifica
'COSTOS VARIABLES
'COSTO FINAL DE SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular la suma de los costos de cotizacion por cantidad de cotizacion del inventario de domingo
miSQL = "SELECT numerosemana([Inventario Cotizacion]![Fecha]) AS semana, Sum(CostoUnitarioUnidad([Inventario Cotizacion]![CodCotizacion],numerosemana([Inventario Cotizacion]![Fecha]))*[Inventario Cotizacion]![Cantidad]) AS Costeo, DBIngredientes.TIPO1 FROM DBIngredientes INNER JOIN ([Inventario Cotizacion] INNER JOIN Cotizaciones ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente GROUP BY numerosemana([Inventario Cotizacion]![Fecha]), DBIngredientes.TIPO1 HAVING (((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem & ") AND ((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoInventarioCotizacion = rst("Costeo")
rst.Close

Exit Function

Nulo:
CostoInventarioCotizacion = 0

End Function

Function CostoInventarioIngrediente(sem As Integer) As Double
'Costo de los items de ingredientes realizados los domingos, de una semana especifica
'COSTOS VARIABLES
'CONSTO INVNTARIO FINAL DE SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular la suma de los costos de ingredientes por cantidad de ingredientes del inventario de domingo
miSQL = "SELECT numerosemana([Inventario Ingrediente]![Fecha]) AS semana, DBIngredientes.TIPO1, Sum(CostoIngredienteGlobal([Inventario Ingrediente]![CodIngrediente],numerosemana([Inventario Ingrediente]![Fecha]))*[Inventario Ingrediente]![Cantidad]) AS Costeo FROM DBIngredientes INNER JOIN [Inventario Ingrediente] ON DBIngredientes.CodIngrediente = [Inventario Ingrediente].CodIngrediente GROUP BY numerosemana([Inventario Ingrediente]![Fecha]), DBIngredientes.TIPO1 HAVING (((numerosemana([Inventario Ingrediente]![Fecha]))=" & sem & ") AND ((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoInventarioIngrediente = rst("Costeo")
rst.Close

Exit Function

Nulo:
CostoInventarioIngrediente = 0

End Function

Function CostoCompras(sem As Integer) As Double
'Costo de los items de Ingresos cotizacion realizados en la semana

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular la suma de los costos de ingresos realizados en la semana
miSQL = "SELECT DBIngredientes.TIPO1, numerosemana([IngresosPitaya]![Fecha]) AS semana, Sum(CostoUnitarioUnidad([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))*[IngresosPitaya]![Cantidad]) AS Costeo FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion GROUP BY DBIngredientes.TIPO1, numerosemana([IngresosPitaya]![Fecha]) HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((numerosemana([IngresosPitaya]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoCompras = rst("Costeo")
rst.Close

Exit Function

Nulo:
CostoCompras = 0

End Function

Function CostoInventarioMain(sem As Integer) As Double
'Costo de los items de cotizacion del invetario principal realizados los domingos, de una semana especifica


On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Calcular la suma de los costos de cotizacion por cantidad de cotizacion del inventario de domingo AMACEN PRINCIPAL
miSQL = "SELECT numerosemana([Inventario Main]![Fecha]) AS semana, Sum(CostoUnitarioUnidad([Inventario Main]![CodCotizacion],numerosemana([Inventario Main]![Fecha]))*[Inventario Main]![Cantidad]) AS Costeo, DBIngredientes.TIPO1 FROM DBIngredientes INNER JOIN ([Inventario Main] INNER JOIN Cotizaciones ON [Inventario Main].CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente GROUP BY numerosemana([Inventario Main]![Fecha]), DBIngredientes.TIPO1 HAVING (((numerosemana([Inventario Main]![Fecha]))=" & sem & ") AND ((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoInventarioMain = rst("Costeo")
rst.Close

Exit Function

Nulo:
CostoInventarioMain = 0

End Function
Function CostoTotalBatido(bat As String, sem As Integer, Cond As Integer) As Double
' cond si es verdadero(-1): vidrio o falso(0): plastico
'costo de batido si es de plastico o vidrio

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Hallar la costo total especifico en la receta especifica
miSQL = "SELECT SubReceta.CodBatido," & _
" Sum([SubReceta]![Cantidad]*" & _
" FactorDeUso([SubReceta]![CodIngrediente],Pedidoplasticovidrio(" & Cond & "))*" & _
" CostoIngredienteGlobal([SubReceta]![CodIngrediente]," & sem & ")) AS Costo" & _
" FROM SubReceta GROUP BY SubReceta.CodBatido HAVING (((SubReceta.CodBatido)='" & bat & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoTotalBatido = rst("Costo")
rst.Close
     
Exit Function

Nulo:
CostoTotalBatido = 0

End Function

Function PedidoPlasticoVidrio(Cond As Integer) As Long
'Busca cualquier nota de pedido que sea plastico o vidrio
'cond:
    '-1: Verdadero osea vidrio
    '0 : Falso osea plastico
'Resultado: Nota de pedido de quesa de esa condicion

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodPedido, Modalidad FROM NotaDePedido WHERE Modalidad=" & Cond & ""
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PedidoPlasticoVidrio = rst("CodPedido")
rst.Close
     
Exit Function

Nulo:
PedidoPlasticoVidrio = PedidoPlasticoVidrio(0)

End Function

Function escotiglobal(coti As Integer) As Long

escotiglobal = DLookup("[Global]", "Cotizaciones", "[CodCotizacion] = " & coti)

End Function

Function CostoSemanalConsumo(sem As Integer) As Long
' Costo de consumo teorico de batido de un ingrediente por su costo unitario, solo VARIABLES

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Verifica si hay consumo, si no es 0, sino suma
miSQL = "SELECT DBIngredientes.TIPO1, Sum(IIf(ConsumoSemanalProducto([DBIngredientes]![CodIngrediente]," & sem & ")=0,0,ConsumoSemanalProducto([DBIngredientes]![CodIngrediente]," & sem & ")*costoingredienteglobal([DBIngredientes]![CodIngrediente]," & sem & "))) AS Monto FROM DBIngredientes GROUP BY DBIngredientes.TIPO1 HAVING (((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoSemanalConsumo = rst("Monto")
rst.Close
     
Exit Function

Nulo:
CostoSemanalConsumo = 0

End Function

Function ValorCotizacionGlobal(coti As Integer, sem As Integer) As Double

On Error GoTo Nulo
Dim Valor As Double

If esporcionablecotizacion(coti) <> 0 Then 'CUando es una porcion
    ValorCotizacionGlobal = DLookup("[Conversion]", "Cotizaciones", "[CodCotizacion] = " & coti) * FRCostoIngredienteGlobal(DLookup("[CodIngrediente]", "Cotizaciones", "[CodCotizacion] = " & coti), sem)
    Exit Function
End If

If codigoLocal() = 0 Then 'Sistema local
    ValorCotizacionGlobal = ValorCotizacionGlobalLocal(coti, sem)
    Exit Function
End If

If DLookup("[Global]", "Cotizaciones", "[CodCotizacion] = " & coti) = 0 Then 'global = 0 'Productos que se compran para local solamente
    Valor = ValorCotizacionGlobalLocal(coti, sem)
    If Valor = 0 Then
        Valor = ValorCotizacionGlobalMixed(coti, sem)
    End If
Else    ' COsto global, que se compra para todos los locales
    Valor = ValorCotizacionGlobalMixed(coti, sem)
End If

ValorCotizacionGlobal = Valor

Exit Function

Nulo:
MsgBox "error al hallar valor de cotizacion"
ValorCotizacionGlobal = 0
End Function

Function ValorCotizacionGlobalLocal(coti As Integer, sem As Integer) As Double
'Costo Unitario promedio de una cotizacion de la semana mencionada hacia abajo EN LOCAL
'SI HAY REGISTRO DE COMPRA LOCAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim stock As Double

' "Costo LOCAL sem" & sem & "coti " & coti
'Calcula el coso unitario de cotizacion promedio en local
miSQL = "SELECT numerosemana([Compras]![Fecha]) AS semana, Compras.CodCotizacion," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Sum(Compras.Cantidad) AS SumaDeCantidad" & _
" FROM Compras" & _
" GROUP BY numerosemana([Compras]![Fecha]), Compras.CodCotizacion" & _
" HAVING (((numerosemana([Compras]![Fecha])) = " & sem & ") And ((Compras.CodCotizacion) = " & coti & ") And" & _
" ((Sum(Compras.CostoTotal)) <> 0) And ((Sum(Compras.Cantidad)) <> 0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
stock = StockSinProcesar(coti, sem)
ValorCotizacionGlobalLocal = (rst("SumaDeCostoTotal") + stock * FRValorCotizacion(coti, sem - 1)) / (rst("SumaDeCantidad") + stock)
rst.Close

Exit Function

Nulo:
ValorCotizacionGlobalLocal = FRValorCotizacion(coti, sem - 1)
End Function

Function ValorCotizacionGlobalMixed(coti As Integer, sem As Integer) As Double
'Costo Unitario promedio de una cotizacion de la semana mencionada hacia abajo EN MIXED
'CUANDO NON HAY VALOR EN BASE DE DATOS LOCAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

' "Costo mix sem" & sem & "coti " & coti
'Calcula el coso unitario de cotizacion promedio en mixed

miSQL = "SELECT numerosemana([Compras]![Fecha]) AS semana, Compras.CodCotizacion," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Sum(Compras.Cantidad) AS SumaDeCantidad, IIf([Compras]![local]=codigolocal(),1,0) AS cond" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY numerosemana([Compras]![Fecha]), Compras.CodCotizacion, IIf([Compras]![local]=codigolocal(),1,0)" & _
" HAVING (((numerosemana([Compras]![Fecha])) = " & sem & ") And ((Compras.CodCotizacion) = " & coti & ") And" & _
" ((Sum(Compras.CostoTotal)) <> 0) And ((Sum(Compras.Cantidad)) <> 0))" & _
" ORDER BY IIf([Compras]![local]=codigolocal(),1,0) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
stock = StockSinProcesar(coti, sem)
ValorCotizacionGlobalMixed = (rst("SumaDeCostoTotal") + stock * FRValorCotizacion(coti, sem - 1)) / (rst("SumaDeCantidad") + stock)
rst.Close

Exit Function

Nulo:
ValorCotizacionGlobalMixed = FRValorCotizacion(coti, sem - 1)
End Function

Function SumaCostosTotalesFactura(numfactura As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma total de factura acorde a número de factura
miSQL = "SELECT Compras.NumeroFactura, Sum(Compras.CostoTotal) AS SumaDeCostoTotal" & _
" FROM Compras GROUP BY Compras.NumeroFactura HAVING (((Compras.NumeroFactura)= '" & numfactura & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaCostosTotalesFactura = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SumaCostosTotalesFactura = 0

End Function

Function SumaCostosTotalesFacturaSucursal(numfactura As String) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma total de factura acorde a número de factura
miSQL = "SELECT ComprasSucursal.NumeroFactura, Sum(ComprasSucursal.CostoTotal) AS SumaDeCostoTotal" & _
" FROM ComprasSucursal GROUP BY ComprasSucursal.NumeroFactura HAVING (((ComprasSucursal.NumeroFactura)= '" & numfactura & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SumaCostosTotalesFacturaSucursal = rst("SumaDeCostoTotal")
rst.Close

Exit Function

Nulo:
SumaCostosTotalesFacturaSucursal = 0

End Function

