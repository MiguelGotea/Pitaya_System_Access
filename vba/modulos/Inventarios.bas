' ==========================================================
' Modulo  : Inventarios
' Tipo    : 1  |  Lineas: 2236
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Function ComprasGlobales(sem As Integer) As Double
'Costo de COmpras globales de la semana, gasto de compras totales para inv pruncipal y locales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las costos de compras totales de la semana VARIABLES
miSQL = "SELECT Sum(ComprasSemana([Cotizaciones]![CodCotizacion]," & sem & ")*CostoUnitarioUnidad([Cotizaciones]![CodCotizacion]," & sem & ")) AS Suma," & _
" DBIngredientes.TIPO1 FROM DBIngredientes INNER JOIN Cotizaciones" & _
" ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.TIPO1 HAVING (((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasGlobales = rst("Suma")
rst.Close

Exit Function

Nulo:
ComprasGlobales = 0

End Function
Function ComprasSemana(coti As Integer, sem As Integer) As Double
'Catidad COmpras totales de un codigo de cotiazCION EN LA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las compras de cierto codigo de cotizacion
miSQL = "SELECT Compras.CodCotizacion, Sum(Compras.Cantidad) AS SumaDeCantidad, numerosemana([Compras]![Fecha]) AS semana" & _
" FROM Compras GROUP BY Compras.CodCotizacion, numerosemana([Compras]![Fecha]) HAVING (((Compras.CodCotizacion)=" & coti & ") AND ((numerosemana([Compras]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasSemana = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ComprasSemana = 0

End Function

Function IngresosSemanaIngredienteBruto(ingre As String, sem As Integer) As Double
'Catidad ingresos totales de un codigo de cotiazCION EN LA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las compras de cierto codigo de cotizacion
miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([IngresosPitaya]![Fecha]) AS semana," & _
" Sum([IngresosPitaya]![Cantidad]*[Cotizaciones]![Conversion]) AS Total" & _
" FROM Cotizaciones INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([IngresosPitaya]![Fecha])" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & ingre & "') AND ((numerosemana([IngresosPitaya]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresosSemanaIngredienteBruto = rst("Total")
rst.Close

Exit Function

Nulo:
IngresosSemanaIngredienteBruto = 0

End Function

Function ComprasSemanaIngredienteBruto(ingre As String, sem As Integer) As Double
'Catidad ingresos totales de un codigo de cotiazCION EN LA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las compras de cierto codigo de cotizacion
miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([Compras]![Fecha]) AS semana," & _
" Sum([Compras]![Cantidad]*[Cotizaciones]![Conversion]) AS Total" & _
" FROM Cotizaciones INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Compras]![Fecha])" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & ingre & "') AND ((numerosemana([Compras]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasSemanaIngredienteBruto = rst("Total")
rst.Close

Exit Function

Nulo:
ComprasSemanaIngredienteBruto = 0

End Function
Function ComprasGlobalesMes(Mes As Integer, ano As Integer) As Double
'Costo de COmpras globales del mes, gasto de compras totales para inv pruncipal y locales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de las costos de compras totales del mes VARIABLES
miSQL = "SELECT DBIngredientes.TIPO1, Month([Compras]![Fecha]) AS Mes, Year([Compras]![Fecha]) AS Año, Sum([Compras]![Cantidad]*CostoUnitarioUnidad([Compras]![CodCotizacion],numerosemana([Compras]![Fecha]))) AS Expr1 FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion GROUP BY DBIngredientes.TIPO1, Month([Compras]![Fecha]), Year([Compras]![Fecha]) HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((Month([Compras]![Fecha]))=" & Mes & ") AND ((Year([Compras]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasGlobalesMes = rst("Expr1")
rst.Close

Exit Function

Nulo:
ComprasGlobalesMes = 0

End Function
Function StockSinProcesar(coti As Integer, sem As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, Sum([Inventario Cotizacion].Cantidad) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS Expr1" & _
" FROM [Inventario Cotizacion]" & _
" GROUP BY [Inventario Cotizacion].CodCotizacion, numerosemana([Inventario Cotizacion]![Fecha])" & _
" HAVING ((([Inventario Cotizacion].CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockSinProcesar = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockSinProcesar = 0

End Function
Function StockSinProcesarUnidadUnitario(coti As Integer, sem As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, Sum([Inventario Cotizacion].cantidadunidad) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS Expr1" & _
" FROM [Inventario Cotizacion]" & _
" GROUP BY [Inventario Cotizacion].CodCotizacion, numerosemana([Inventario Cotizacion]![Fecha])" & _
" HAVING ((([Inventario Cotizacion].CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockSinProcesarUnidadUnitario = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockSinProcesarUnidadUnitario = 0

End Function
Function StockSinProcesarUnidadPaquete(coti As Integer, sem As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, Sum([Inventario Cotizacion].cantidadpaquete) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS Expr1" & _
" FROM [Inventario Cotizacion]" & _
" GROUP BY [Inventario Cotizacion].CodCotizacion, numerosemana([Inventario Cotizacion]![Fecha])" & _
" HAVING ((([Inventario Cotizacion].CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockSinProcesarUnidadPaquete = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockSinProcesarUnidadPaquete = 0

End Function

Function AjustesSinProcesar(coti As Integer, sem As Integer) As Double
'ajustes de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT AjustesInventario.CodCotizacion, Sum(AjustesInventario.Cantidad) AS SumaDeCantidad," & _
" numerosemana(AjustesInventario.Fecha) AS Expr1" & _
" FROM AjustesInventario" & _
" GROUP BY AjustesInventario.CodCotizacion, numerosemana(AjustesInventario.Fecha)" & _
" HAVING (((AjustesInventario.CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana(AjustesInventario.Fecha))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AjustesSinProcesar = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
AjustesSinProcesar = 0

End Function

Function StockSinProcesarInventarioTemporal(coti As Integer, sem As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion Temporal].CodCotizacion, Sum([Inventario Cotizacion Temporal].Cantidad) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion Temporal]![Fecha]) AS Expr1" & _
" FROM [Inventario Cotizacion Temporal]" & _
" GROUP BY [Inventario Cotizacion Temporal].CodCotizacion, numerosemana([Inventario Cotizacion Temporal]![Fecha])" & _
" HAVING ((([Inventario Cotizacion Temporal].CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([Inventario Cotizacion Temporal]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockSinProcesarInventarioTemporal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockSinProcesarInventarioTemporal = 0

End Function
Function StockCantidadUnidadInventarioTemporal(coti As Integer, sem As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion Temporal].CodCotizacion, Sum([Inventario Cotizacion Temporal].cantidadunidad) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion Temporal]![Fecha]) AS Expr1" & _
" FROM [Inventario Cotizacion Temporal]" & _
" GROUP BY [Inventario Cotizacion Temporal].CodCotizacion, numerosemana([Inventario Cotizacion Temporal]![Fecha])" & _
" HAVING ((([Inventario Cotizacion Temporal].CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([Inventario Cotizacion Temporal]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCantidadUnidadInventarioTemporal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockCantidadUnidadInventarioTemporal = 0

End Function
Function StockCantidadPaqueteInventarioTemporal(coti As Integer, sem As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion Temporal].CodCotizacion, Sum([Inventario Cotizacion Temporal].cantidadpaquete) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion Temporal]![Fecha]) AS Expr1" & _
" FROM [Inventario Cotizacion Temporal]" & _
" GROUP BY [Inventario Cotizacion Temporal].CodCotizacion, numerosemana([Inventario Cotizacion Temporal]![Fecha])" & _
" HAVING ((([Inventario Cotizacion Temporal].CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([Inventario Cotizacion Temporal]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCantidadPaqueteInventarioTemporal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockCantidadPaqueteInventarioTemporal = 0

End Function

Function StockSinProcesarxLocal(coti As Integer, sem As Integer, loc As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial de local

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad inventariada de cierta cotizacion
miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, Sum([Inventario Cotizacion].Cantidad) AS SumaDeCantidad," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS Expr1, [Inventario Cotizacion].local" & _
" FROM [Inventario Cotizacion] IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY [Inventario Cotizacion].CodCotizacion, numerosemana([Inventario Cotizacion]![Fecha]), [Inventario Cotizacion].local" & _
" HAVING ([Inventario Cotizacion].CodCotizacion=" & coti & "" & _
" AND numerosemana([Inventario Cotizacion]![Fecha])=" & sem - 1 & "" & _
" AND [Inventario Cotizacion].local=" & loc & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockSinProcesarxLocal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockSinProcesarxLocal = 0

End Function

Function StockSinProcesarxLocalxPorductoVenta(bati As String, sem As Integer, loc As Integer) As Double
'Stock de productops cotizacion Sin Procesar de ñla semana, stock inicial de local

On Error GoTo Nulo

Dim rst2 As DAO.Recordset
Dim miSQL2 As String

Dim canti As Double
Dim baticod As String
Dim cotibatix As Long

Dim rst As DAO.Recordset
Dim miSQL As String


StockSinProcesarxLocalxPorductoVenta = 0

miSQL2 = "SELECT DBBatidos.Nombre, DBBatidos.Vigencia, CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido]) AS cotibati" & _
" FROM DBBatidos GROUP BY DBBatidos.Nombre, DBBatidos.Vigencia, CotiPrincipalProdCompraVenta([DBBatidos]![CodBatido])" & _
" HAVING (((DBBatidos.Nombre)='" & bati & "') AND ((DBBatidos.Vigencia)<>0))"
Set rst2 = CurrentDb.OpenRecordset(miSQL2, dbOpenDynaset)
rst2.MoveLast
canti = rst2.RecordCount
rst2.MoveFirst

For I = 1 To canti
    cotibatix = rst2("cotibati")

    'SUmma de la cantidad inventariada de cierta cotizacion
    miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, Sum([Inventario Cotizacion].Cantidad) AS SumaDeCantidad," & _
    " numerosemana([Inventario Cotizacion]![Fecha]) AS Expr1, [Inventario Cotizacion].local" & _
    " FROM [Inventario Cotizacion] IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
    " GROUP BY [Inventario Cotizacion].CodCotizacion, numerosemana([Inventario Cotizacion]![Fecha]), [Inventario Cotizacion].local" & _
    " HAVING ([Inventario Cotizacion].CodCotizacion=" & cotibatix & "" & _
    " AND numerosemana([Inventario Cotizacion]![Fecha])=" & sem - 1 & "" & _
    " AND [Inventario Cotizacion].local=" & loc & ")"
    Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
    StockSinProcesarxLocalxPorductoVenta = StockSinProcesarxLocalxPorductoVenta + rst("SumaDeCantidad")
    rst.Close


    rst2.MoveNext
Next I

rst2.Close

Exit Function

Nulo:
StockSinProcesarxLocalxPorductoVenta = 0

End Function
Function StockSinProcesarIngrediente(ING As String, sem As Integer) As Double
'Stock de productops Sin Procesar de ñla semana, suma de cotizaciones de mismo ingrediente

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad cotizaciones por su conversion natural de un ingreidnete
miSQL = "SELECT Sum(StockSinProcesar([Cotizaciones]![CodCotizacion]," & sem & ")*ConversionNatural([Cotizaciones]![CodCotizacion])) AS Suma, Cotizaciones.CodIngrediente FROM Cotizaciones GROUP BY Cotizaciones.CodIngrediente HAVING (((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockSinProcesarIngrediente = rst("Suma")
rst.Close

Exit Function

Nulo:
StockSinProcesarIngrediente = 0

End Function
Function IngresosPitaya(coti As Integer, sem As Integer) As Double
'Cantidad Total de cotizacion Ingresos a Local de cierta cotizacion en una semana especifica
'CANTIDAD DE UNIDADES DE COTIZACION UNGRESADAS EN UNA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad ingresos de cierta coptoizacion
miSQL = "SELECT IngresosPitaya.CodCotizacion, Sum(IngresosPitaya.Cantidad) AS SumaDeCantidad," & _
" numerosemana([IngresosPitaya]![Fecha]) AS Expr1" & _
" FROM IngresosPitaya GROUP BY IngresosPitaya.CodCotizacion, numerosemana([IngresosPitaya]![Fecha])" & _
" HAVING (((IngresosPitaya.CodCotizacion)=" & coti & ") AND ((numerosemana([IngresosPitaya]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresosPitaya = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
IngresosPitaya = 0

End Function

Function IngresosPitayaRango(coti As Integer, semd As Integer, semh As Integer) As Double
'Cantidad Total de cotizacion Ingresos a Local de cierta cotizacion en una semana especifica
'CANTIDAD DE UNIDADES DE COTIZACION UNGRESADAS EN UNA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad ingresos de cierta coptoizacion
miSQL = "SELECT IngresosPitaya.CodCotizacion, Sum(IngresosPitaya.Cantidad) AS SumaDeCantidad," & _
" numerosemana([IngresosPitaya]![Fecha])<=" & semh & " AND numerosemana([IngresosPitaya]![Fecha])>=" & semd & " AS condi" & _
" FROM IngresosPitaya" & _
" GROUP BY IngresosPitaya.CodCotizacion," & _
" numerosemana([IngresosPitaya]![Fecha])<=" & semh & " AND numerosemana([IngresosPitaya]![Fecha])>=" & semd & "" & _
" HAVING (((IngresosPitaya.CodCotizacion)=" & coti & ")" & _
" AND ((numerosemana([IngresosPitaya]![Fecha])<=" & semh & " AND numerosemana([IngresosPitaya]![Fecha])>=" & semd & ")<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresosPitayaRango = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
IngresosPitayaRango = 0

End Function
Function IngresosPitayaIngrediente(ING As String, sem As Integer) As Double
'Cantidad Total de Ingresos a Local de cierta ingrediente proviene de suma de cotizaciones natural
'CANTIDAD DE INGREDIENTE INGRESADO , INGREDIENTE YA PROCESADO

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cotizaciones a nivel de ingrediente ingresados en una semana especifica a pitaya
miSQL = "SELECT Sum(IngresosPitaya([Cotizaciones]![CodCotizacion]," & sem & ")*ConversionNatural([Cotizaciones]![CodCotizacion])) AS Suma, Cotizaciones.CodIngrediente FROM Cotizaciones GROUP BY Cotizaciones.CodIngrediente HAVING (((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresosPitayaIngrediente = rst("Suma")
rst.Close

Exit Function

Nulo:
IngresosPitayaIngrediente = 0

End Function
Function TotalProcesado(coti As Integer, sem As Integer) As Double
'Cantidad Total de cotizacion propcesados de una cotizacionb especifica
'CANTIDAD DE PRODUCTOS PROCESADOS EN UNA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUmma de la cantidad unidades procesados
miSQL = "SELECT Procesamiento.CodCotizacion, Sum(Procesamiento.Cantidad) AS SumaDeCantidad," & _
" numerosemana([Procesamiento]![Fecha]) AS Expr1" & _
" FROM Procesamiento GROUP BY Procesamiento.CodCotizacion, numerosemana([Procesamiento]![Fecha])" & _
" HAVING (((Procesamiento.CodCotizacion)=" & coti & ") AND ((numerosemana([Procesamiento]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TotalProcesado = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
TotalProcesado = 0

End Function
Function ProcesadoIngrediente(ING As String, semana As Integer) As Long
'Cantidada procesada de un ingrediente especifico en una semana especifica
'CANTIDAD DE INGREDIENTE PROCESADO EN UNA SEMANA

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
'SUma total de lo procesado de un ingredente
miSQL = "SELECT Sum(Procesamiento.MedidaFinal) AS SumaDeMedidaFinal, Cotizaciones.CodIngrediente, numerosemana([Procesamiento]![Fecha]) AS Semana FROM Procesamiento INNER JOIN Cotizaciones ON Procesamiento.CodCotizacion = Cotizaciones.CodCotizacion GROUP BY Cotizaciones.CodIngrediente, numerosemana([Procesamiento]![Fecha]) HAVING (((Cotizaciones.CodIngrediente)='" & ING & "') AND ((numerosemana([Procesamiento]![Fecha]))=" & semana & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ProcesadoIngrediente = rst("SumaDeMedidaFinal")
rst.Close

Exit Function

Nulo:
ProcesadoIngrediente = 0

End Function



Function StockCotizacion(inge As String, sem As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingrediente
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente," & _
" Sum([Inventario Cotizacion]![Cantidad]*conversionestandar([Inventario Cotizacion]![CodCotizacion])) AS total," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS semana" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha])" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacion = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacion = 0

End Function
Function StockCotizacionSinPorciones(inge As String, sem As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingrediente
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente," & _
" Sum([Inventario Cotizacion]![Cantidad]*conversionestandar([Inventario Cotizacion]![CodCotizacion])) AS total," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS semana, Cotizaciones.Subproducto" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]), Cotizaciones.Subproducto" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & ") AND ((Cotizaciones.Subproducto)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacionSinPorciones = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacionSinPorciones = 0

End Function

Function AjustesCotizacion(inge As String, sem As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingrediente
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente," & _
" Sum(AjustesInventario.Cantidad*conversionestandar(AjustesInventario.CodCotizacion)) AS total," & _
" numerosemana(AjustesInventario.Fecha) AS semana" & _
" FROM AjustesInventario INNER JOIN Cotizaciones ON AjustesInventario.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana(AjustesInventario.Fecha)" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana(AjustesInventario.Fecha))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AjustesCotizacion = rst("total")
rst.Close

Exit Function

Nulo:
AjustesCotizacion = 0

End Function
Function AjustesCotizacionSinPorciones(inge As String, sem As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingrediente
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente," & _
" Sum(AjustesInventario.Cantidad*conversionestandar(AjustesInventario.CodCotizacion)) AS total," & _
" numerosemana(AjustesInventario.Fecha) AS semana, Cotizaciones.Subproducto" & _
" FROM AjustesInventario INNER JOIN Cotizaciones ON AjustesInventario.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana(AjustesInventario.Fecha), Cotizaciones.Subproducto" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana(AjustesInventario.Fecha))=" & sem & ") AND ((Cotizaciones.Subproducto)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
AjustesCotizacionSinPorciones = rst("total")
rst.Close

Exit Function

Nulo:
AjustesCotizacionSinPorciones = 0

End Function

Function StockCotizacionProductosSinConversion(ING As String, sem As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingrediente
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente, Cotizaciones.Conversion," & _
" Sum([Inventario Cotizacion]![Cantidad]) AS total," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS semana" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, Cotizaciones.Conversion, numerosemana([Inventario Cotizacion]![Fecha])" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & ING & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & ")" & _
" AND ((Cotizaciones.Conversion)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacionProductosSinConversion = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacionProductosSinConversion = 0

End Function

Function StockCotizacionlocalconversionestandar(inge As String, sem As Integer, loc As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingredientede loca especidifco
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]) AS semana," & _
" [Inventario Cotizacion].local," & _
" Sum([Inventario Cotizacion]![Cantidad]*conversionestandar([Inventario Cotizacion]![CodCotizacion])) AS total" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones" & _
" ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]), [Inventario Cotizacion].local" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & ")" & _
" AND (([Inventario Cotizacion].local)=" & loc & "));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacionlocalconversionestandar = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacionlocalconversionestandar = 0

End Function

Function StockCotizacionlocal(inge As String, sem As Integer, loc As Integer) As Double

'Stock acorde a cantidad de cotizaciones valor resultado en medida ingredientede loca especidifco
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]) AS semana," & _
" [Inventario Cotizacion].local," & _
" Sum([Inventario Cotizacion]![Cantidad]*conversionestandar([Inventario Cotizacion]![CodCotizacion])) AS total" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones" & _
" ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]), [Inventario Cotizacion].local" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem - 1 & ")" & _
" AND (([Inventario Cotizacion].local)=" & loc & "));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacionlocal = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacionlocal = 0

End Function

Function StockFinalSoloPorcioneslocal(inge As String, sem As Integer, loc As Integer) As Double
'Stock acorde a cantidad de cotizaciones valor resultado en medida ingredientede loca especidifco, solo prociones
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente , solo porciones subproducto <>0
miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]) AS semana," & _
" [Inventario Cotizacion].local," & _
" Sum([Inventario Cotizacion]![Cantidad]*conversionestandar([Inventario Cotizacion]![CodCotizacion])) AS total," & _
" [Cotizaciones].Subproducto" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones" & _
" ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]), [Inventario Cotizacion].local, [Cotizaciones].Subproducto" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & inge & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))=" & sem & ")" & _
" AND (([Inventario Cotizacion].local)=" & loc & ") AND ((Cotizaciones.Subproducto)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockFinalSoloPorcioneslocal = rst("total")
rst.Close

Exit Function

Nulo:
StockFinalSoloPorcioneslocal = 0
End Function

Function StockCotizacionConversionNatural(ING As String, sem As Integer) As Double

'Stock acorde a cantidad de cotizaciones conversion natural
'stock inicial de ingrediente en base a cotizacion es stock final de una semana anterior sem-1

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'cantidad de stock de cierto ingrediente acorde a cantidad de cotizaiones que hay de ese ingrediente
miSQL = "SELECT Cotizaciones.CodIngrediente," & _
" Sum([Inventario Cotizacion]![Cantidad]*ConversionCalculado([Inventario Cotizacion]![CodCotizacion], " & sem - 1 & ")) AS total," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS semana, [Cotizaciones]![Conversion]<>0 AS [natural]" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]), [Cotizaciones]![Conversion]<>0" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & ING & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))= " & sem - 1 & ")" & _
" AND (([Cotizaciones]![Conversion]<>0)<>0));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacionConversionNatural = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacionConversionNatural = 0

End Function

Function StockCotizacionPorcionable(ingre As String, sem As Integer) As Double
'SI
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.CodIngrediente," & _
" Sum([Inventario Cotizacion]![Cantidad]*ConversionCalculado([Inventario Cotizacion]![CodCotizacion], " & sem - 1 & ")) AS total," & _
" numerosemana([Inventario Cotizacion]![Fecha]) AS semana, [Cotizaciones]![Conversion]<>0 AS [natural], Cotizaciones.Subproducto" & _
" FROM [Inventario Cotizacion] INNER JOIN Cotizaciones ON [Inventario Cotizacion].CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.CodIngrediente, numerosemana([Inventario Cotizacion]![Fecha]), [Cotizaciones]![Conversion]<>0, Cotizaciones.Subproducto" & _
" HAVING (((Cotizaciones.CodIngrediente)='" & ingre & "') AND ((numerosemana([Inventario Cotizacion]![Fecha]))= " & sem - 1 & ") AND (([Cotizaciones]![Conversion]<>0)<>0) AND ((Cotizaciones.Subproducto)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockCotizacionPorcionable = rst("total")
rst.Close

Exit Function

Nulo:
StockCotizacionPorcionable = 0

End Function

Function StockIngrediente(ingex As String, sem As Integer) As Double
' stock inicial de ingrediente es stock final de una semana anterior sem-1
'Stock de ingrediente procesado
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'setock de ingrediente en bruto
miSQL = "SELECT numerosemana([Inventario Ingrediente]![Fecha]) AS semana," & _
" [Inventario Ingrediente].CodIngrediente," & _
" Sum([Inventario Ingrediente].Cantidad) AS SumaDeCantidad" & _
" FROM [Inventario Ingrediente]" & _
" GROUP BY numerosemana([Inventario Ingrediente]![Fecha]), [Inventario Ingrediente].CodIngrediente" & _
" HAVING (((numerosemana([Inventario Ingrediente]![Fecha]))=" & sem - 1 & ") AND (([Inventario Ingrediente].CodIngrediente)='" & ingex & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockIngrediente = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockIngrediente = 0

End Function

Function StockIngredienteTemporal(ING As String, sem As Integer) As Double
' stock inicial de ingrediente es stock final de una semana anterior sem-1
'Stock de ingrediente procesado
'De la tabla temporal
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'setock de ingrediente en bruto
miSQL = "SELECT numerosemana([Inventario Ingrediente Temporal]![Fecha]) AS semana," & _
" [Inventario Ingrediente Temporal].CodIngrediente," & _
" Sum([Inventario Ingrediente Temporal].Cantidad) AS SumaDeCantidad" & _
" FROM [Inventario Ingrediente Temporal]" & _
" GROUP BY numerosemana([Inventario Ingrediente Temporal]![Fecha]), [Inventario Ingrediente Temporal].CodIngrediente" & _
" HAVING (((numerosemana([Inventario Ingrediente Temporal]![Fecha]))=" & sem - 1 & ") AND (([Inventario Ingrediente Temporal].CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockIngredienteTemporal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockIngredienteTemporal = 0

End Function
Function StockIngredientelocal(inge As String, sem As Integer, loc As Integer) As Double
' stock inicial de ingrediente es stock final de una semana anterior sem-1 de basedatos global
'Stock de ingrediente procesado
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'setock de ingrediente en bruto
miSQL = "SELECT numerosemana([Inventario Ingrediente]![Fecha]) AS semana," & _
" [Inventario Ingrediente].CodIngrediente, [Inventario Ingrediente].local," & _
" Sum([Inventario Ingrediente].Cantidad) AS SumaDeCantidad" & _
" FROM [Inventario Ingrediente] IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY numerosemana([Inventario Ingrediente]![Fecha]), [Inventario Ingrediente].CodIngrediente," & _
" [Inventario Ingrediente].local" & _
" HAVING (((numerosemana([Inventario Ingrediente]![Fecha]))=" & sem - 1 & ")" & _
" AND (([Inventario Ingrediente].CodIngrediente)='" & inge & "') AND (([Inventario Ingrediente].local)=" & loc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockIngredientelocal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockIngredientelocal = 0

End Function


Function ConsumoSemanalProducto(ingref As String, semana As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de consumo en semana especifica de ingrediente
miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Semana," & _
" Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Consumo," & _
" SubReceta.CodIngrediente, NotaDePedido.Anulado" & _
" FROM (DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" ON DBBatidos.CodBatido = SubPedido.CodBatido) INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), SubReceta.CodIngrediente, NotaDePedido.Anulado" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & semana & ")" & _
" AND ((SubReceta.CodIngrediente)='" & ingref & "') AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoSemanalProducto = rst("Consumo")
rst.Close

Exit Function

Nulo:
ConsumoSemanalProducto = 0

End Function

Function ConsumoSemanalMaximoIngrediente(inge As String, semana As Integer, rango As Integer) As Double

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de consumo maximo en semana especifica de ingrediente
miSQL = "SELECT TOP 1 NotaDePedido.Anulado, SubReceta.CodIngrediente, numerosemana([NotaDePedido]![Fecha]) AS Semana," & _
" Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Consumo" & _
" FROM (DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" ON DBBatidos.CodBatido = SubPedido.CodBatido)" & _
" INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, SubReceta.CodIngrediente, numerosemana([NotaDePedido]![Fecha])" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((SubReceta.CodIngrediente)='" & inge & "')" & _
" AND ((numerosemana([NotaDePedido]![Fecha])) Between " & semana - rango & " And " & semana - 1 & "))" & _
" ORDER BY Sum([SubPedido]![Cantidad]*[SubReceta]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoSemanalMaximoIngrediente = rst("Consumo")
rst.Close

Exit Function

Nulo:
ConsumoSemanalMaximoIngrediente = 0

End Function

Function CalVolumenConsumoSemanal(sem As Integer) As Double  'Local
'Suma de gramos de insumos y de productos no en gramos transformados a gramos

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim suma As Double
suma = 0

miSQL = "SELECT TIPO1, TIPO2, Tipo, CodIngrediente, ConversionGramos" & _
" FROM DBIngredientes WHERE (TIPO1='VARIABLES' AND TIPO2='Batidos' AND Tipo<>'Empaque')"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

Do While Not rst.EOF
    suma = suma + FRConsumoIngrediente(rst("CodIngrediente"), sem) * rst("ConversionGramos")
    rst.MoveNext
Loop

rst.Close

CalVolumenConsumoSemanal = suma

Exit Function

Nulo:
CalVolumenConsumoSemanal = 0

End Function
Function IngresoTotalIngrediente(ING As String, semana As Integer) As Long
'CANTIDAD DE INGREDIENTE INGRESADO NATURAL Y POR CONVERSION

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de INGREDIENTE ingresado en una semana natural y por conversion
miSQL = "SELECT numerosemana([IngresosPitaya]![Fecha]) AS semana," & _
" Sum([IngresosPitaya]![Cantidad]*ConversionCalculado([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS Cantidad," & _
" Cotizaciones.CodIngrediente" & _
" FROM IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([IngresosPitaya]![Fecha]), Cotizaciones.CodIngrediente" & _
" HAVING (((numerosemana(IngresosPitaya!Fecha))=" & semana & ") And ((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresoTotalIngrediente = rst("Cantidad")
rst.Close

Exit Function

Nulo:
IngresoTotalIngrediente = 0

End Function

Function IngresoTotalIngredienteSinPorciones(ingex As String, semana As Integer) As Long
'CANTIDAD DE INGREDIENTE INGRESADO NATURAL Y POR CONVERSION

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de INGREDIENTE ingresado en una semana natural y por conversion
miSQL = "SELECT numerosemana([IngresosPitaya]![Fecha]) AS semana," & _
" Sum([IngresosPitaya]![Cantidad]*conversionestandar([IngresosPitaya]![CodCotizacion])) AS Cantidad," & _
" Cotizaciones.CodIngrediente, Cotizaciones.SubProducto, [Cotizaciones]![Marca] & ' '='Almacen Global ' AS almacenglobal" & _
" FROM IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([IngresosPitaya]![Fecha]), Cotizaciones.CodIngrediente," & _
" Cotizaciones.SubProducto, [Cotizaciones]![Marca] & ' '='Almacen Global '" & _
" HAVING (((numerosemana(IngresosPitaya!Fecha))=" & semana & ") And ((Cotizaciones.CodIngrediente)='" & ingex & "')" & _
" AND ((Cotizaciones.Subproducto)=0) AND (([Cotizaciones]![Marca] & ' '='Almacen Global ')=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresoTotalIngredienteSinPorciones = rst("Cantidad")
rst.Close

Exit Function

Nulo:
IngresoTotalIngredienteSinPorciones = 0

End Function

Function IngresoTotalIngredienteConversionNatural(ING As String, semana As Integer) As Long
'CANTIDAD DE INGREDIENTE INGRESADO NATURAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de INGREDIENTE ingresado en una semana natural y por conversion
miSQL = "SELECT numerosemana([IngresosPitaya]![Fecha]) AS semana, Cotizaciones.CodIngrediente," & _
" [Cotizaciones]![Conversion]<>0 AS [natural]," & _
" Sum([IngresosPitaya]![Cantidad]*ConversionCalculado([IngresosPitaya]![CodCotizacion]," & semana & ")) AS Cantidad" & _
" FROM IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([IngresosPitaya]![Fecha]), Cotizaciones.CodIngrediente, [Cotizaciones]![Conversion]<>0" & _
" HAVING (((numerosemana([IngresosPitaya]![Fecha]))=" & semana & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "') AND (([Cotizaciones]![Conversion]<>0)<>0));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresoTotalIngredienteConversionNatural = rst("Cantidad")
rst.Close

Exit Function

Nulo:
IngresoTotalIngredienteConversionNatural = 0

End Function

Function IngresoTotalIngredienteSinConversion(ING As String, semana As Integer) As Long
'CANTIDAD DE INGREDIENTE INGRESADO NATURAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Cantidad de INGREDIENTE ingresado en una semana natural y por conversion
miSQL = "SELECT numerosemana([IngresosPitaya]![Fecha]) AS semana, Cotizaciones.CodIngrediente," & _
" [Cotizaciones]![Conversion]=0 AS [natural]," & _
" Sum([IngresosPitaya]![Cantidad]) AS Cantidad" & _
" FROM IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY numerosemana([IngresosPitaya]![Fecha]), Cotizaciones.CodIngrediente, [Cotizaciones]![Conversion]=0" & _
" HAVING (((numerosemana([IngresosPitaya]![Fecha]))=" & semana & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "') AND (([Cotizaciones]![Conversion]=0)<>0));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresoTotalIngredienteSinConversion = rst("Cantidad")
rst.Close

Exit Function

Nulo:
IngresoTotalIngredienteSinConversion = 0

End Function

Function consumoreal(ING As String, semana As Integer) As Long
'CONSUMO REAL DE UN INGREDIENTE

On Error GoTo Nulo

consumoreal = StockCotizacion(ING, semana) + StockIngrediente(ING, semana) + IngresoTotalIngrediente(ING, semana) - StockCotizacion(ING, semana + 1) - StockIngrediente(ING, semana + 1) - MermaIngrediente(ING, semana) - MermaCotizacion(ING, semana)
Exit Function

Nulo:
consumoreal = 0

End Function

Function MermaCotizacion(ING As String, sem As Integer) As Double

'Cantidad de merma convertido a valor ingrediente de una semana especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de cotizaciones por conversion de un ingrediente en una semana
miSQL = "SELECT Sum([Merma Cotizacion]![Cantidad]*ConversionCalculado([Merma Cotizacion]![CodCotizacion],numerosemana([Merma Cotizacion]![Fecha]))) AS total, numerosemana([Merma Cotizacion]![Fecha]) AS semana, Cotizaciones.CodIngrediente FROM Cotizaciones INNER JOIN [Merma Cotizacion] ON Cotizaciones.CodCotizacion = [Merma Cotizacion].CodCotizacion GROUP BY numerosemana([Merma Cotizacion]![Fecha]), Cotizaciones.CodIngrediente HAVING (((numerosemana([Merma Cotizacion]![Fecha]))=" & sem & ") AND ((Cotizaciones.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MermaCotizacion = rst("total")
rst.Close

Exit Function

Nulo:
MermaCotizacion = 0

End Function

Function MermaCotizacionSinPorciones(ingre As String, sem As Integer) As Double

'Cantidad de merma convertido a valor ingrediente de una semana especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de cotizaciones por conversion de un ingrediente en una semana
miSQL = "SELECT Sum([Merma Cotizacion]![Cantidad]*ConversionCalculado([Merma Cotizacion]![CodCotizacion]," & _
" numerosemana([Merma Cotizacion]![Fecha]))) AS total, numerosemana([Merma Cotizacion]![Fecha]) AS semana," & _
" Cotizaciones.CodIngrediente, Cotizaciones.SubProducto" & _
" FROM Cotizaciones" & _
" INNER JOIN [Merma Cotizacion]" & _
" ON Cotizaciones.CodCotizacion = [Merma Cotizacion].CodCotizacion" & _
" GROUP BY numerosemana([Merma Cotizacion]![Fecha]), Cotizaciones.CodIngrediente, Cotizaciones.SubProducto" & _
" HAVING (((numerosemana([Merma Cotizacion]![Fecha]))=" & sem & ")" & _
" AND ((Cotizaciones.CodIngrediente)='" & ingre & "') AND ((Cotizaciones.SubProducto)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MermaCotizacionSinPorciones = rst("total")
rst.Close

Exit Function

Nulo:
MermaCotizacionSinPorciones = 0

End Function

Function MermaIngrediente(ingrex As String, sem As Integer) As Double
'Cantidad de ingrediente en valor ingrediente de una semana

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'suma de montos de ingrediente merma en una semana
miSQL = "SELECT numerosemana([Merma Unidad]![Fecha]) AS semana, [Merma Unidad].CodIngrediente," & _
" Sum([Merma Unidad].Cantidad) AS SumaDeCantidad" & _
" FROM [Merma Unidad]" & _
" GROUP BY numerosemana([Merma Unidad]![Fecha]), [Merma Unidad].CodIngrediente" & _
" HAVING (((numerosemana([Merma Unidad]![Fecha]))=" & sem & ") AND (([Merma Unidad].CodIngrediente)='" & ingrex & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MermaIngrediente = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
MermaIngrediente = 0

End Function

Function MermaUnidadesCotizacion(coti As Integer, sem As Integer) As Double
'Cantidad de cotizaciones en valor cotizacion de una semana

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'suma de montos de cotizaciones merma en una semana
miSQL = "SELECT numerosemana([Merma Cotizacion]![Fecha]) AS semana, [Merma Cotizacion].CodCotizacion," & _
" Sum([Merma Cotizacion].Cantidad) AS SumaDeCantidad FROM [Merma Cotizacion] GROUP BY numerosemana([Merma Cotizacion]![Fecha])," & _
" [Merma Cotizacion].CodCotizacion HAVING (((numerosemana([Merma Cotizacion]![Fecha]))=" & sem & ")" & _
" AND (([Merma Cotizacion].CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MermaUnidadesCotizacion = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
MermaUnidadesCotizacion = 0

End Function
Function MermaUnidadesCotizacionRango(coti As Integer, semd As Integer, semh As Integer) As Double
'Cantidad de cotizaciones en valor cotizacion de una semana

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'suma de montos de cotizaciones merma en una semana
miSQL = "SELECT numerosemana([Merma Cotizacion]![Fecha])>=" & semd & " And numerosemana([Merma Cotizacion]![Fecha])<=" & semh & " AS semana," & _
" [Merma Cotizacion].CodCotizacion, Sum([Merma Cotizacion].Cantidad) AS SumaDeCantidad" & _
" FROM [Merma Cotizacion]" & _
" GROUP BY numerosemana([Merma Cotizacion]![Fecha])>=" & semd & " And numerosemana([Merma Cotizacion]![Fecha])<=" & semh & "," & _
" [Merma Cotizacion].CodCotizacion" & _
" HAVING (((numerosemana([Merma Cotizacion]![Fecha])>=" & semd & " And numerosemana([Merma Cotizacion]![Fecha])<=" & semh & ")<>0)" & _
" AND (([Merma Cotizacion].CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
MermaUnidadesCotizacionRango = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
MermaUnidadesCotizacionRango = 0

End Function
Function porcionadoprocesado(cod As Integer) As Long
'SUma de pesos porcionados total, se necesita codigo de procesamiento

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'COnversion por la cantidad de porciones ingresadas bajo un codigo de porcionamiento
miSQL = "SELECT Sum([Porcionamiento]![Cantidad]*[Cotizaciones]![Conversion]) AS total," & _
" Porcionamiento.CodProcesamiento" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Porcionamiento.CodProcesamiento HAVING (((Porcionamiento.CodProcesamiento)=" & cod & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
porcionadoprocesado = rst("total")
rst.Close

Exit Function

Nulo:
porcionadoprocesado = 0

End Function
Function porcionadoselladoxsubpor(cod As Integer) As Long
'cod es codigo de subporcionamiento

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'COnversion por la cantidad de porciones ingresadas bajo un codigo de porcionamiento
miSQL = "SELECT Sum([Porcionamiento]![Cantidad]*[Cotizaciones]![Conversion]) AS total, Porcionamiento.CodSubPorcionamiento" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Porcionamiento.CodSubPorcionamiento" & _
" HAVING (((Porcionamiento.CodSubPorcionamiento)=" & cod & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
porcionadoselladoxsubpor = rst("total")
rst.Close

Exit Function

Nulo:
porcionadoselladoxsubpor = 0

End Function
Function porcionadoselladoxsubporxlocal(cod As Integer, loca As Integer) As Long
'cod es codigo de subporcionamiento

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'COnversion por la cantidad de porciones ingresadas bajo un codigo de porcionamiento
miSQL = "SELECT Sum([Porcionamiento" & loca & "]![Cantidad]*[Cotizaciones]![Conversion]) AS total, Porcionamiento" & loca & ".CodSubPorcionamiento" & _
" FROM Cotizaciones INNER JOIN Porcionamiento" & loca & " ON Cotizaciones.CodCotizacion = Porcionamiento" & loca & ".CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Porcionamiento" & loca & ".CodSubPorcionamiento" & _
" HAVING (((Porcionamiento" & loca & ".CodSubPorcionamiento)=" & cod & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
porcionadoselladoxsubporxlocal = rst("total")
rst.Close

Exit Function

Nulo:
porcionadoselladoxsubporxlocal = 0

End Function
Function porcionadoenteros(proce As Integer, sem As Integer) As Long
'SUma de pesos porcionados total de una cotizacion porcionada

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'COnversion por la cantidad de porciones ingresadas bajo un codigo de cotizacion y una semana especifica
miSQL = "SELECT Sum([Porcionamiento]![Cantidad]*[Cotizaciones]![Conversion]) AS total, Porcionamiento.CodProcesamiento," & _
" numerosemana([Porcionamiento]![Fecha]) AS semana, Porcionamiento.Procedencia" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Porcionamiento.CodProcesamiento, numerosemana([Porcionamiento]![Fecha]), Porcionamiento.Procedencia" & _
" HAVING (((Porcionamiento.CodProcesamiento)=0) AND ((numerosemana([Porcionamiento]![Fecha]))=" & sem & ")" & _
" AND ((Porcionamiento.Procedencia)=" & proce & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
porcionadoenteros = rst("total")
rst.Close

Exit Function

Nulo:
porcionadoenteros = 0

End Function

Function ingresoporcionessemana(coti As Integer, sem As Integer) As Long
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de procion especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Porcionamiento.CodCotizacion, numerosemana([Porcionamiento]![Fecha]) AS semana," & _
" Sum(Porcionamiento.Cantidad) AS SumaDeCantidad" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Porcionamiento.CodCotizacion, numerosemana([Porcionamiento]![Fecha])" & _
" HAVING (Porcionamiento.CodCotizacion=" & coti & " AND numerosemana([Porcionamiento]![Fecha])=" & sem & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ingresoporcionessemana = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ingresoporcionessemana = 0

End Function

Function ingresoporcionessemanarango(coti As Integer, semd As Integer, semh As Integer) As Double
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de procion especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Porcionamiento.CodCotizacion," & _
" numerosemana([Porcionamiento]![Fecha])<=" & semh & " And numerosemana([Porcionamiento]![Fecha])>=" & semd & " AS semana," & _
" Sum(Porcionamiento.Cantidad) AS SumaDeCantidad" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Porcionamiento.CodCotizacion," & _
" numerosemana([Porcionamiento]![Fecha])<=" & semh & " And numerosemana([Porcionamiento]![Fecha])>=" & semd & "" & _
" HAVING (((Porcionamiento.CodCotizacion)=" & coti & ") AND" & _
" ((numerosemana([Porcionamiento]![Fecha])<=" & semh & " And numerosemana([Porcionamiento]![Fecha])>=" & semd & ")<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ingresoporcionessemanarango = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ingresoporcionessemanarango = 0

End Function

Function ingresoporcionessemanaxdespacho(coti As Integer, sem As Integer) As Long
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de procion especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Porcionamiento.Procedencia, Porcionamiento.CodCotizacion, numerosemana([Porcionamiento]![Fecha]) AS semana," & _
" Sum(Porcionamiento.Cantidad) AS SumaDeCantidad" & _
" FROM Porcionamiento" & _
" GROUP BY Porcionamiento.Procedencia, Porcionamiento.CodCotizacion, numerosemana([Porcionamiento]![Fecha])" & _
" HAVING (((Porcionamiento.Procedencia)=PorcionGlobalDePorcion([Porcionamiento]![CodCotizacion]))" & _
" AND ((Porcionamiento.CodCotizacion)=" & coti & ") AND ((numerosemana([Porcionamiento]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ingresoporcionessemanaxdespacho = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ingresoporcionessemanaxdespacho = 0

End Function

Function ingresoporcionessemanaxtransformacion(coti As Integer, sem As Integer) As Long
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de procion especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Porcionamiento.Procedencia, Porcionamiento.CodCotizacion, numerosemana([Porcionamiento]![Fecha]) AS semana," & _
" Sum(Porcionamiento.Cantidad) AS SumaDeCantidad" & _
" FROM Porcionamiento" & _
" GROUP BY Porcionamiento.Procedencia, Porcionamiento.CodCotizacion, numerosemana([Porcionamiento]![Fecha])" & _
" HAVING (((Porcionamiento.Procedencia)<>PorcionGlobalDePorcion([Porcionamiento]![CodCotizacion]))" & _
" AND ((Porcionamiento.CodCotizacion)=" & coti & ") AND ((numerosemana([Porcionamiento]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ingresoporcionessemanaxtransformacion = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ingresoporcionessemanaxtransformacion = 0

End Function
Function ingresoingredientesinporcionessemanaxtransformacion(ingrex As String, sem As Integer) As Long
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de procion especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Cotizaciones.Subproducto, Cotizaciones.CodIngrediente, numerosemana([Porcionamiento]![Fecha]) AS semana," & _
" Sum([Porcionamiento]![Cantidad]*ConversionEstandar([Porcionamiento]![CodCotizacion])) AS Cantidad" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY Cotizaciones.Subproducto, Cotizaciones.CodIngrediente, numerosemana([Porcionamiento]![Fecha])" & _
" HAVING (((Cotizaciones.Subproducto)=0) AND ((Cotizaciones.CodIngrediente)='" & ingrex & "')" & _
" AND ((numerosemana([Porcionamiento]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ingresoingredientesinporcionessemanaxtransformacion = rst("Cantidad")
rst.Close

Exit Function

Nulo:
ingresoingredientesinporcionessemanaxtransformacion = 0

End Function

Function ingresoporcionamientodeconversioncerosemana(ingex As String, sem As Integer) As Long
'Cantidad de ingrediente que ingreso por porcionamiento de producto con conversion=0

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [Porcionamiento]![Procedencia])=0 AS [natural]," & _
" Cotizaciones.CodIngrediente, numerosemana([Porcionamiento]![Fecha]) AS semana," & _
" Sum([Porcionamiento]![Cantidad]*[Cotizaciones]![Conversion]) AS total" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" GROUP BY DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [Porcionamiento]![Procedencia])=0," & _
" Cotizaciones.CodIngrediente, numerosemana([Porcionamiento]![Fecha])" & _
" HAVING (((DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [Porcionamiento]![Procedencia])=0)<>0)" & _
" AND ((Cotizaciones.CodIngrediente)='" & ingex & "') AND ((numerosemana([Porcionamiento]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ingresoporcionamientodeconversioncerosemana = rst("total")
rst.Close

Exit Function

Nulo:
ingresoporcionamientodeconversioncerosemana = 0

End Function

Function cotizacionesporcionadas(coti As Integer, sem As Integer) As Double
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de cotizacion especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Porcionamiento.Procedencia, numerosemana([Porcionamiento]![Fecha]) AS semana, Cotizaciones.Conversion," & _
" Sum([Porcionamiento]![Cantidad]*ConversionCalculado([Porcionamiento]![CodCotizacion]," & sem & ")/[Cotizaciones]![Conversion]) AS Expr1" & _
" FROM Porcionamiento INNER JOIN Cotizaciones ON Porcionamiento.Procedencia = Cotizaciones.CodCotizacion" & _
" GROUP BY Porcionamiento.Procedencia, numerosemana([Porcionamiento]![Fecha]), Cotizaciones.Conversion" & _
" HAVING (Porcionamiento.Procedencia=" & coti & " AND numerosemana([Porcionamiento]![Fecha])=" & sem & "" & _
" AND Cotizaciones.Conversion<>0)"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cotizacionesporcionadas = rst("Expr1")
rst.Close

Exit Function

Nulo:
cotizacionesporcionadas = 0

End Function
Function cotizacionesporcionadasreales(coti As Integer, sem As Integer) As Double
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de cotizacion especifica
'sacado de registro de subporcionamiento

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPorcionamiento.Procedencia, numerosemana([SubPorcionamiento]![Fecha]) AS sema," & _
" Sum(SubPorcionamiento.Cantidad) AS SumaDeCantidad FROM SubPorcionamiento" & _
" GROUP BY SubPorcionamiento.Procedencia, numerosemana([SubPorcionamiento]![Fecha])" & _
" HAVING (((SubPorcionamiento.Procedencia)=" & coti & ") AND ((numerosemana([SubPorcionamiento]![Fecha]))=" & sem & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cotizacionesporcionadasreales = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cotizacionesporcionadasreales = 0

End Function

Function cotizacionesporcionadasrealesrango(coti As Integer, semd As Integer, semh As Integer) As Double
'Cantidad de items porcionados en la semana especifica, codigo de cotizacion de cotizacion especifica
'sacado de registro de subporcionamiento

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT SubPorcionamiento.Procedencia," & _
" numerosemana([SubPorcionamiento]![Fecha])>=" & semd & " And numerosemana([SubPorcionamiento]![Fecha])<=" & semh & " AS sema," & _
" Sum(SubPorcionamiento.Cantidad) AS SumaDeCantidad" & _
" FROM SubPorcionamiento" & _
" GROUP BY SubPorcionamiento.Procedencia," & _
" numerosemana([SubPorcionamiento]![Fecha])>=" & semd & " And numerosemana([SubPorcionamiento]![Fecha])<=" & semh & "" & _
" HAVING (((SubPorcionamiento.Procedencia)=" & coti & ")" & _
" AND ((numerosemana([SubPorcionamiento]![Fecha])>=" & semd & " And numerosemana([SubPorcionamiento]![Fecha])<=" & semh & ")<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cotizacionesporcionadasrealesrango = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
cotizacionesporcionadasrealesrango = 0

End Function
Function consumoporciones(coti As Integer, sem As Integer) As Long
'Cantidad cotizaciones de porcion usadas en una semana
'coti: porcion predeinida subprod<>0

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Expr1, SubReceta.codporcion, NotaDePedido.Anulado," & _
" Sum([SubReceta]![Cantidad]/DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])*[SubPedido]![Cantidad]) AS CantPorc" & _
" FROM (NotaDePedido" & _
" INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN (SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido)" & _
" ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), SubReceta.codporcion, NotaDePedido.Anulado" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha]))=" & sem & ") AND ((SubReceta.codporcion)=" & coti & ") AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumoporciones = rst("CantPorc")
rst.Close

Exit Function

Nulo:
consumoporciones = 0

End Function

Function consumomaximoporciones(coti As Integer, sem As Integer, rang As Integer) As Double
'Cantidad cotizaciones de porcion usadas en una semana de las ultimas rang semanas
'coti: porcion predeinida subprod<>0

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([NotaDePedido]![Fecha]) AS Expr1, SubReceta.codporcion, NotaDePedido.Anulado," & _
" Sum([SubReceta]![Cantidad]/DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])*[SubPedido]![Cantidad]) AS CantPorc" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN (SubReceta INNER JOIN DBBatidos ON SubReceta.CodBatido = DBBatidos.CodBatido) ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY numerosemana([NotaDePedido]![Fecha]), SubReceta.codporcion, NotaDePedido.Anulado" & _
" HAVING (((numerosemana([NotaDePedido]![Fecha])) Between " & sem - rang + 1 & " And " & sem & ")" & _
" AND ((SubReceta.codporcion)=" & coti & ") AND ((NotaDePedido.Anulado)=0))" & _
" ORDER BY Sum([SubReceta]![Cantidad]/DLookUp('[Conversion]','[Cotizaciones]','[CodCotizacion]=' & [SubReceta]![codporcion])*[SubPedido]![Cantidad]) DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumomaximoporciones = rst("CantPorc")
rst.Close

Exit Function

Nulo:
consumomaximoporciones = 0

End Function

Function consumoingredientexporciones(ainge As String, sem As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente," & _
" Sum(consumoporciones([Cotizaciones]![CodCotizacion]," & sem & ")*[Cotizaciones]![Conversion]) AS Total" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.CodIngrediente" & _
" HAVING (((DBIngredientes.CodIngrediente)='" & ainge & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumoingredientexporciones = rst("Total")
rst.Close

Exit Function

Nulo:
consumoingredientexporciones = 0

End Function

Function consumomaximoingredientexporcioneslocal(ainge As String, sem As Integer, locx As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, de un local especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente," & _
" Sum(FRConsumoMaximoPorcionxLocal([Cotizaciones]![CodCotizacion]," & sem & ", " & locx & ")*[Cotizaciones]![Conversion]) AS Total" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.CodIngrediente" & _
" HAVING (((DBIngredientes.CodIngrediente)='" & ainge & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumomaximoingredientexporcioneslocal = rst("Total")
rst.Close

Exit Function

Nulo:
consumomaximoingredientexporcioneslocal = 0

End Function

Function consumoingredientexporcioneslocal(ainge As String, sem As Integer, locx As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, de un local especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente," & _
" Sum(FRConsumoPorcionxLocal([Cotizaciones]![CodCotizacion]," & sem & ", " & locx & ")*[Cotizaciones]![Conversion]) AS Total" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.CodIngrediente" & _
" HAVING (((DBIngredientes.CodIngrediente)='" & ainge & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumoingredientexporcioneslocal = rst("Total")
rst.Close

Exit Function

Nulo:
consumoingredientexporcioneslocal = 0

End Function

Function FRconsumoingredientexporciones(ainge As String, sem As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, de un local especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente," & _
" Sum(FRConsumoPorcion([Cotizaciones]![CodCotizacion]," & sem & ")*[Cotizaciones]![Conversion]) AS Total" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.CodIngrediente" & _
" HAVING (((DBIngredientes.CodIngrediente)='" & ainge & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRconsumoingredientexporciones = rst("Total")
rst.Close

Exit Function

Nulo:
FRconsumoingredientexporciones = 0

End Function

Function FRconsumoingredientexporcionesSumaLocales(ingresl As String, semasl As Integer) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
miSQL = "SELECT StatusSucursales.CodLocal, StatusSucursales.Activo, StatusSucursales.Sucursal" & _
" FROM StatusSucursales WHERE ((StatusSucursales.Sucursal<>0) AND (StatusSucursales.Activo<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
FRconsumoingredientexporcionesSumaLocales = 0
Do While Not rst.EOF

    FRconsumoingredientexporcionesSumaLocales = FRconsumoingredientexporcionesSumaLocales + FRconsumoingredientexporciones(ingresl, semasl)
    rst.MoveNext
Loop
rst.Close
Exit Function

Nulo:
FRconsumoingredientexporcionesSumaLocales = 0
End Function
Function consumomaximoingredientexporcionesxDespacholocal(ainge As String, sem As Integer, locx As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, de un local especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente," & _
" Sum((redondear_mas(FRConsumoMaximoPorcionxLocal([Cotizaciones]![CodCotizacion]," & sem & "," & locx & ")/IIf(IsNull([Cotizaciones]![PaquetePorciones]),1,[Cotizaciones]![PaquetePorciones]))*IIf(IsNull([Cotizaciones]![PaquetePorciones]),1,[Cotizaciones]![PaquetePorciones]))*[Cotizaciones]![Conversion]) AS Total" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.CodIngrediente" & _
" HAVING (((DBIngredientes.CodIngrediente)='" & ainge & "'))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumomaximoingredientexporcionesxDespacholocal = rst("Total")
rst.Close

Exit Function

Nulo:
consumomaximoingredientexporcionesxDespacholocal = 0

End Function
Function consumoingredientexporcionesxDespacholocal(ainge As String, sem As Integer, locx As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, de un local especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT DBIngredientes.CodIngrediente," & _
" Sum((redondear_mas(FRConsumoPorcionxLocal([Cotizaciones]![CodCotizacion]," & sem & "," & locx & ")/IIf(IsNull([Cotizaciones]![PaquetePorciones]),1,[Cotizaciones]![PaquetePorciones]))*IIf(IsNull([Cotizaciones]![PaquetePorciones]),1,[Cotizaciones]![PaquetePorciones]))*[Cotizaciones]![Conversion]) AS Total" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY DBIngredientes.CodIngrediente" & _
" HAVING (((DBIngredientes.CodIngrediente)='" & ainge & "'))"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
consumoingredientexporcionesxDespacholocal = rst("Total")
rst.Close

Exit Function

Nulo:
consumoingredientexporcionesxDespacholocal = 0

End Function

Function consumoingredientexporcionesmaximo(ainge As String, sem As Integer, ran As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel
'maximo de las ultimas ran semanas desde la semana antes hacia atras
Dim tempok As Double
Dim maxi As Double
maxi = 0

For I = 1 To ran
    tempok = consumoingredientexporciones(ainge, sem - I)
    If tempok > maxi Then
        maxi = tempok
    End If
Next I
consumoingredientexporcionesmaximo = maxi

Exit Function

Nulo:
consumoingredientexporcionesmaximo = 0

End Function

Function consumoingredientexporcionesmaximolocalcalculado(ainge As String, sem As Integer, ran As Integer, locas As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, en un local especifico
'maximo de las ultimas ran semanas desde la semana antes hacia atras, en un lcoal especifico

If ran = 4 Then 'aja directo de la db
    consumoingredientexporcionesmaximolocalcalculado = consumomaximoingredientexporcioneslocal(ainge, sem - 1, locas)
Else
    consumoingredientexporcionesmaximolocalcalculado = consumoingredientexporcionesmaximolocal(ainge, sem, ran, locas)
End If

Exit Function

Nulo:
consumoingredientexporcionesmaximolocalcalculado = 0

End Function

Function consumoingredientexporcionesmaximolocal(ainge As String, sem As Integer, ran As Integer, locas As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, en un local especifico
'maximo de las ultimas ran semanas desde la semana antes hacia atras, en un lcoal especifico
Dim tempok As Double
Dim maxi As Double
maxi = 0

For I = 1 To ran
    tempok = consumoingredientexporcioneslocal(ainge, sem - 1, locas)
    If tempok > maxi Then
        maxi = tempok
    End If
Next I
consumoingredientexporcionesmaximolocal = maxi

Exit Function

Nulo:
consumoingredientexporcionesmaximolocal = 0

End Function
Function consumoingredientexporcionesmaximoxDespacholocalCalculado(ainge As String, sem As Integer, ran As Integer, locas As Integer) As Double

'Cantidad ingrediente usados pero como porciones ,no a granel, en un local especifico
'maximo de las ultimas ran semanas desde la semana antes hacia atras, en un lcoal especifico
On Error GoTo Nulo

If ran = 4 Then ' por defecto si es 4 jala de db
    consumoingredientexporcionesmaximoxDespacholocalCalculado = consumomaximoingredientexporcionesxDespacholocal(ainge, sem - 1, locas)
Else
    consumoingredientexporcionesmaximoxDespacholocalCalculado = consumoingredientexporcionesmaximoxDespacholocal(ainge, sem, ran, locas)
End If

Exit Function

Nulo:
consumoingredientexporcionesmaximoxDespacholocalCalculado = 0

End Function

Function consumoingredientexporcionesmaximoxDespacholocal(ainge As String, sem As Integer, ran As Integer, locas As Integer) As Double
'Cantidad ingrediente usados pero como porciones ,no a granel, en un local especifico
'maximo de las ultimas ran semanas desde la semana antes hacia atras, en un lcoal especifico
Dim tempok As Double
Dim maxi As Double

maxi = 0


For I = 1 To ran
    tempok = consumoingredientexporcionesxDespacholocal(ainge, sem - I, locas)
    If tempok > maxi Then
        maxi = tempok
    End If
Next I
consumoingredientexporcionesmaximoxDespacholocal = maxi

Exit Function

Nulo:
consumoingredientexporcionesmaximoxDespacholocal = 0

End Function
Function TipoDeGrupo(grup As Long) As String

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
miSQL = "SELECT Grupos.Tipo, Grupos.CodGrupo FROM Grupos WHERE (Grupos.CodGrupo=" & grup & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TipoDeGrupo = rst("Tipo")
rst.Close
Exit Function
Nulo:
TipoDeGrupo = ""
End Function
Function StockFinalMain(coti As Integer, sem As Integer) As Double

'Stock FInal de Almacen principal

StockFinalMain = StockSinProcesar(coti, sem + 1)

Exit Function

Nulo:
StockFinalMain = 0

End Function

Function ComprasMain(coti As Integer, sem As Integer) As Double
'compras de un producto todos los lcoales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodCotizacion, Sum(Cantidad) AS SumaDeCantidad," & _
" numerosemana(Fecha) AS semana" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY CodCotizacion, numerosemana(Fecha)" & _
" HAVING (CodCotizacion=" & coti & " AND numerosemana(Fecha)=" & sem & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasMain = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ComprasMain = 0

End Function

Function EgresosMain(coti As Integer, sem As Integer) As Double
'Salidas a los locales desde almacen principal de un producto todos los lcoales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodCotizacion, Sum(Cantidad) AS SumaDeCantidad," & _
" numerosemana(Fecha) AS semana" & _
" FROM IngresosPitaya IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY CodCotizacion, numerosemana(Fecha)" & _
" HAVING (CodCotizacion=" & coti & " AND numerosemana(Fecha)=" & sem & ")"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
EgresosMain = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
EgresosMain = 0

End Function

Function EgresosMesMain(codc As Integer, Mes As Integer, ano As Integer) As Long
'Ingresos totales de productos cotizacion  en el mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de cantidades ingresadas de cierto producto en mes especifico por local " & pi & " de ano especifico
miSQL = "SELECT IngresosPitaya.CodCotizacion, Sum(IngresosPitaya.Cantidad) AS SumaDeCantidad," & _
" Month([IngresosPitaya]![Fecha]) AS Mes, Year([IngresosPitaya]![Fecha]) AS Año" & _
" FROM IngresosPitaya IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY IngresosPitaya.CodCotizacion, Month([IngresosPitaya]![Fecha]), Year([IngresosPitaya]![Fecha])" & _
" HAVING (((IngresosPitaya.CodCotizacion)=" & codc & ") AND ((Month([IngresosPitaya]![Fecha]))=" & Mes & ")" & _
" AND ((Year([IngresosPitaya]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
EgresosMesMain = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
EgresosMesMain = 0

End Function



Function ComprasMesMain(codc As Integer, Mes As Integer, ano As Integer) As Long
'Compras de productos cotizacion  en el mes especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de cantidades compradas de cierto producto en mes especifico por local " & pi & " de ano especifico
miSQL = "SELECT Compras.CodCotizacion, Month([Compras]![Fecha]) AS Mes, Year([Compras]![Fecha]) AS Año," & _
" Sum(Compras.Cantidad) AS SumaDeCantidad" & _
" FROM Compras IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Compras.CodCotizacion, Month([Compras]![Fecha]), Year([Compras]![Fecha])" & _
" HAVING (((Compras.CodCotizacion)=" & codc & ") AND ((Month([Compras]![Fecha]))=" & Mes & ")" & _
" AND ((Year([Compras]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasMesMain = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
ComprasMesMain = 0

End Function

Function StockFinalMesMain(codc As Integer, Mes As Integer, ano As Integer) As Double
'stock final  de productos cotizacion  en el mes especifico, es el stock final del mes presente
' el stock se toma fin de mes por lo que el inicial del mes es el tomado el mes anterior

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [Inventario Main].CodCotizacion, Month([Inventario Main]![Fecha]) AS Mes," & _
" Year([Inventario Main]![Fecha]) AS Año, Sum([Inventario Main].Cantidad) AS SumaDeCantidad" & _
" FROM [Inventario Main]" & _
" GROUP BY [Inventario Main].CodCotizacion, Month([Inventario Main]![Fecha]), Year([Inventario Main]![Fecha])" & _
" HAVING ((([Inventario Main].CodCotizacion)=" & codc & ") AND ((Month([Inventario Main]![Fecha]))=" & Mes & ")" & _
" AND ((Year([Inventario Main]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
StockFinalMesMain = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
StockFinalMesMain = 0

End Function

Function StockFinalTeoricoPorciones(codc As Integer, sem As Integer) As Double

On Error GoTo Nulo

StockFinalTeoricoPorciones = StockSinProcesar(codc, sem) + ingresoporcionessemana(codc, sem) - consumoporciones(codc, sem)

Exit Function

Nulo:
StockFinalTeoricoPorciones = 0

End Function

Function StockFinalTeoricoCompraVenta(codb As String, sem As Integer) As Double

On Error GoTo Nulo
Dim cotipv As Integer
Dim ingrepv As String
Dim nomb As String
Dim medi As String
Dim tipopv As String
nomb = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & codb & "'")
medi = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & codb & "'")
cotipv = CotiPrincipalProdCompraVenta(codb)
ingrepv = IngrePrincipalProdCompraVenta(codb)

tipopv = IIf(TipoOrigenProductoCompraVenta(codb) = 0, "porcion", "ingrediente")

If tipopv = "porcion" Then
    StockFinalTeoricoCompraVenta = StockSinProcesar(cotipv, sem) + ingresoporcionessemana(cotipv, sem) - MermaUnidadesCotizacion(cotipv, sem) - ventaproductosemana(nomb, medi, sem)
Else
    StockFinalTeoricoCompraVenta = StockIngrediente(ingrepv, sem) + StockCotizacion(ingrepv, sem) + IngresoTotalIngrediente(ingrepv, sem) - (MermaIngrediente(ingrepv, sem) + MermaCotizacion(ingrepv, sem)) - ventaproductosemana(nomb, medi, sem)
End If
Exit Function
    
Nulo:
StockFinalTeoricoCompraVenta = 0

End Function

Function PreIngresoCotizacionSemanaLocal(coti As Integer, sem As Integer, loca As Integer) As Double
'preingreso de sistema 0 a los locales

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT PreIngresoPitaya.Destino, numerosemana([PreIngresoPitaya]![Fecha]) AS sema," & _
" SubPreIngresosPitaya.CodCotizacion, Sum(SubPreIngresosPitaya.Cantidad) AS SumaDeCantidad" & _
" FROM SubPreIngresosPitaya INNER JOIN PreIngresoPitaya" & _
" ON SubPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya" & _
" GROUP BY PreIngresoPitaya.Destino, numerosemana([PreIngresoPitaya]![Fecha]), SubPreIngresosPitaya.CodCotizacion" & _
" HAVING (((PreIngresoPitaya.Destino)='" & "Pitaya " & loca & "') AND ((numerosemana([PreIngresoPitaya]![Fecha]))=" & sem & ")" & _
" AND ((SubPreIngresosPitaya.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PreIngresoCotizacionSemanaLocal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
PreIngresoCotizacionSemanaLocal = 0

End Function

Function PreIngresoCotizacionSemanaTotal(coti As Integer, sem As Integer) As Double
'preingreso de sistema 0 a los todos losloclees

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([PreIngresoPitaya]![Fecha]) AS sema," & _
" SubPreIngresosPitaya.CodCotizacion, Sum(SubPreIngresosPitaya.Cantidad) AS SumaDeCantidad" & _
" FROM SubPreIngresosPitaya INNER JOIN PreIngresoPitaya" & _
" ON SubPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya" & _
" GROUP BY numerosemana([PreIngresoPitaya]![Fecha]), SubPreIngresosPitaya.CodCotizacion" & _
" HAVING (((numerosemana([PreIngresoPitaya]![Fecha]))=" & sem & ")" & _
" AND ((SubPreIngresosPitaya.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PreIngresoCotizacionSemanaTotal = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
PreIngresoCotizacionSemanaTotal = 0

End Function

Function PreIngresoCotizacionSemanaTotalRango(coti As Integer, semd As Integer, semh As Integer) As Double
'preingreso de sistema 0 a los todos losloclees

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([PreIngresoPitaya]![Fecha])>=" & semd & " And numerosemana([PreIngresoPitaya]![Fecha])<=" & semh & " AS sema," & _
" SubPreIngresosPitaya.CodCotizacion, Sum(SubPreIngresosPitaya.Cantidad) AS SumaDeCantidad" & _
" FROM SubPreIngresosPitaya" & _
" INNER JOIN PreIngresoPitaya ON SubPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya" & _
" GROUP BY numerosemana([PreIngresoPitaya]![Fecha])>=" & semd & " And numerosemana([PreIngresoPitaya]![Fecha])<=" & semh & "," & _
" SubPreIngresosPitaya.CodCotizacion" & _
" HAVING (((numerosemana([PreIngresoPitaya]![Fecha])>=" & semd & " And numerosemana([PreIngresoPitaya]![Fecha])<=" & semh & ")<>0)" & _
" AND ((SubPreIngresosPitaya.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
PreIngresoCotizacionSemanaTotalRango = rst("SumaDeCantidad")
rst.Close

Exit Function

Nulo:
PreIngresoCotizacionSemanaTotalRango = 0

End Function


Function StockFinalDiaKardexSemana(inge As String, dial As Date) As Double
'suma de kardex semanal al finalizar un dia

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim sema As Integer
Dim canti As Integer
sema = numerosemana(dial)

miSQL = "SELECT Cotizaciones.CodIngrediente, numerosemana([IngresosPitaya]![Fecha]) AS sem, IngresosPitaya.Fecha," & _
" [IngresosPitaya]![Cantidad]*[Cotizaciones]![Conversion] AS cant" & _
" FROM IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion" & _
" WHERE (((Cotizaciones.CodIngrediente) = '" & inge & "')" & _
" And ((numerosemana([IngresosPitaya]![Fecha])) = " & sema & ")" & _
" And ((IngresosPitaya.Fecha) <= #" & dial & "#))" & _
" UNION ALL SELECT Cotizaciones.CodIngrediente, numerosemana([Procesamiento]![Fecha]) AS sem, Procesamiento.Fecha," & _
" [Procesamiento]![MedidaFinal] AS cant" & _
" FROM Procesamiento INNER JOIN Cotizaciones ON Procesamiento.CodCotizacion = Cotizaciones.CodCotizacion" & _
" WHERE (((Cotizaciones.CodIngrediente) = '" & inge & "')" & _
" And ((numerosemana([Procesamiento]![Fecha])) = " & sema & ")" & _
" And ((Procesamiento.Fecha) <= #" & dial & "#))" & _
" UNION ALL SELECT SubReceta.CodIngrediente, numerosemana([NotaDePedido]![Fecha]) AS sem," & _
" NotaDePedido.Fecha, [SubReceta]![Cantidad]*[SubPedido]![Cantidad]*(-1) AS cant" & _
" FROM ((SubPedido INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido)" & _
" INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) INNER JOIN NotaDePedido" & _
" ON SubPedido.CodPedido = NotaDePedido.CodPedido" & _
" WHERE (((SubReceta.CodIngrediente)='" & inge & "')" & _
" AND ((numerosemana([NotaDePedido]![Fecha]))=" & sema & ")" & _
" AND ((NotaDePedido.Fecha)<=#" & dial & "#) AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
canti = rst.RecordCount
rst.MoveFirst
StockFinalDiaKardexSemana = StockIngrediente(inge, sema) + StockCotizacionConversionNatural(inge, sema)

For I = 1 To canti
    StockFinalDiaKardexSemana = StockFinalDiaKardexSemana + rst("cant")
    rst.MoveNext
Next I
rst.Close

Exit Function

Nulo:
StockFinalDiaKardexSemana = 0

End Function
Function esproductovigentesegunfiltronoporcion(coti As Integer) As Integer
'1: positivo, 0 negativo
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT ListaNOPORCIONESFiltro.CodCotizacion, 1 AS Estado" & _
" FROM ListaNOPORCIONESFiltro" & _
" WHERE (((ListaNOPORCIONESFiltro.CodCotizacion)=" & coti & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
esproductovigentesegunfiltronoporcion = rst("Estado")
rst.Close

Exit Function

Nulo:
esproductovigentesegunfiltronoporcion = 0

End Function

Function esproductovigenteseguninventarioguardado(coti As Integer, fechi As Date, lista As Integer) As Integer
'1: positivo, 0 negativo
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT [Inventario Cotizacion].CodCotizacion, [Inventario Cotizacion].Fecha, [Inventario Cotizacion].lista, 1 AS Estado" & _
" FROM [Inventario Cotizacion]" & _
" WHERE ((([Inventario Cotizacion].CodCotizacion)=" & coti & ")" & _
" AND (([Inventario Cotizacion].Fecha)=#" & fechi & "#)" & _
" AND (([Inventario Cotizacion].lista)=" & lista & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
esproductovigenteseguninventarioguardado = rst("Estado")
rst.Close

Exit Function

Nulo:
esproductovigenteseguninventarioguardado = 0

End Function

Function factorinventariounidadunitario(cotix As Integer) As Double

Dim Subproducto As Integer
Dim Marca As String
Dim PaquetePorciones As Double
Dim Conversion As Double
Dim Unidad As String

Subproducto = DLookup("[subproducto]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
Marca = DLookup("[Marca] & ' '", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
PaquetePorciones = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
Conversion = DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
Unidad = DLookup("[Unidad]", "[DBIngredientes]", "[CodIngrediente]='" & DLookup("[CodIngrediente]", "[Cotizaciones]", "[CodCotizacion]=" & cotix) & "'")

If Subproducto <> 0 Then
    If Marca = "Marca Pitaya " Then
        If PaquetePorciones <> 1 Then
            '"unidad"
            factorinventariounidadunitario = 1
        Else
            '" "
            factorinventariounidadunitario = 0
        End If
    Else
        '"unidad"
        factorinventariounidadunitario = 1
    End If
Else
    If Conversion <> 1 Then
        If Conversion = 0 Then
            '" "
            factorinventariounidadunitario = 0
        Else
            If PaquetePorciones = 1 Then
                If Unidad = "gr" Then
                    '"oz"
                    factorinventariounidadunitario = (28.375) / Conversion
                Else
                    If Unidad = "ramas" Then
                        '" "
                        factorinventariounidadunitario = 0
                    Else
                        '[DBIngredientes]![Unidad]
                        factorinventariounidadunitario = 1 / Conversion
                    End If
                End If
            Else
                '[Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]
                factorinventariounidadunitario = 1
            End If
        End If
    Else
        If PaquetePorciones = 1 Then
            '" "
            factorinventariounidadunitario = 0
        Else
            '[Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]
            factorinventariounidadunitario = 1
        End If
    End If
End If
End Function

Function factorinventariounidaddespacho(cotix As Integer) As Double
Dim Subproducto As Integer
Dim Marca As String
Dim PaquetePorciones As Double
Dim Conversion As Double
Dim Unidad As String

Subproducto = DLookup("[subproducto]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
Marca = DLookup("[Marca] & ' '", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
PaquetePorciones = DLookup("[PaquetePorciones]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
Conversion = DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & cotix)
Unidad = DLookup("[Unidad]", "[DBIngredientes]", "[CodIngrediente]='" & DLookup("[CodIngrediente]", "[Cotizaciones]", "[CodCotizacion]=" & cotix) & "'")

If PaquetePorciones = 1 Then
    If Subproducto <> 0 Then
        '"unidad"
        factorinventariounidaddespacho = 1
    Else
        '[Cotizaciones]![Unidad] & " " & [Cotizaciones]![Capacidad]
        factorinventariounidaddespacho = 1
    End If
Else
    'Paq x " & [Cotizaciones]![PaquetePorciones]
    factorinventariounidaddespacho = PaquetePorciones
End If
End Function
