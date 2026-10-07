' ==========================================================
' Modulo  : ER y FC
' Tipo    : 1  |  Lineas: 343
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:12
' ==========================================================

Option Compare Database
'MES Y AÑO DE FORMULARIO ABIERTO ( FLUJO DE CAJA O ESTADO DE RESULTADOS)


'INGRESOS
Function IngresoMes(Mes As Integer, ano As Integer) As Double
'iNGRESOS en todo el mes del ano especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de monto de venta en mes y ano especifico
miSQL = "SELECT Month([NotaDePedido]![Fecha]) AS Mes, Year([NotaDePedido]![Fecha]) AS ano, Sum(MontoPedido([NotaDePedido]![CodPedido])) AS Monto FROM NotaDePedido GROUP BY Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha]) HAVING (((Month([NotaDePedido]![Fecha]))=" & Mes & ") AND ((Year([NotaDePedido]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresoMes = rst("Monto")
rst.Close

Exit Function

Nulo:
IngresoMes = 0

End Function

Function IngresoMesTeorico(mesi As Integer, ano As Integer) As Double
'Precio real  en todo el mes del ano especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de monto de precio real [por productos en mes y ano especifico
miSQL = "SELECT NotaDePedido.Anulado, Month([NotaDePedido]![Fecha]) AS mesi, Year([NotaDePedido]![Fecha]) AS ano, Sum([SubPedido]![Cantidad]*[DBBatidos]![Precio]) AS Monto" & _
" FROM DBBatidos INNER JOIN (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) ON DBBatidos.CodBatido = SubPedido.CodBatido" & _
" GROUP BY NotaDePedido.Anulado, Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha])" & _
" HAVING (((NotaDePedido.Anulado)=0) AND ((Month([NotaDePedido]![Fecha]))=" & mesi & ") AND ((Year([NotaDePedido]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresoMesTeorico = rst("Monto")
rst.Close

Exit Function

Nulo:
IngresoMesTeorico = 0

End Function

Function IngresosPorNombreGrupo(Tipo As String, Mes As Integer, ano As Integer) As Double
'Monto total de ventas por nombre de grupo enu mes y ano especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de monto de nombre de grupo de un mes y ano especifico
miSQL = "SELECT Grupos.NombreGrupo, Sum([SubPedido]![Cantidad]) AS Monto, Month([NotaDePedido]![Fecha]) AS Mes, Year([NotaDePedido]![Fecha]) AS Ano FROM Grupos INNER JOIN ((NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido) INNER JOIN DBBatidos ON SubPedido.CodBatido = DBBatidos.CodBatido) ON Grupos.CodGrupo = DBBatidos.CodGrupo GROUP BY Grupos.NombreGrupo, Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha]) HAVING (((Grupos.NombreGrupo)='" & Tipo & "') AND ((Month([NotaDePedido]![Fecha]))=" & Mes & ") AND ((Year([NotaDePedido]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
IngresosPorNombreGrupo = rst("Monto")
rst.Close

Exit Function

Nulo:
IngresosPorNombreGrupo = 0

End Function

'EGRESOS
    'METODO 1: Costo de consumo teorico de ingredientes en el mes

Function ConsumoIngredienteMes(inge As String, mesi As Integer, ano As Integer) As Double  'Local
'Cantidad de un codigo de ingrediente consumido en un mes y ano especifico
'costo de venta teorico(CANTIDADES)

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de las cantidades teoricas consumidas de un ingediente especifico en ano y mes dado
miSQL = "SELECT Month([NotaDePedido]![Fecha]) AS mesi, Year([NotaDePedido]![Fecha]) AS ano, SubReceta.CodIngrediente," & _
" NotaDePedido.Anulado," & _
" Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Cantidad" & _
" FROM (SubPedido INNER JOIN (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido) ON SubPedido.CodBatido = DBBatidos.CodBatido) INNER JOIN NotaDePedido ON SubPedido.CodPedido = NotaDePedido.CodPedido" & _
" GROUP BY Month([NotaDePedido]![Fecha]), Year([NotaDePedido]![Fecha]), SubReceta.CodIngrediente, NotaDePedido.Anulado" & _
" HAVING (((Month([NotaDePedido]![Fecha]))=" & mesi & ") AND ((Year([NotaDePedido]![Fecha]))=" & ano & ") AND ((SubReceta.CodIngrediente)='" & inge & "') AND ((NotaDePedido.Anulado)=0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ConsumoIngredienteMes = rst("Cantidad")
rst.Close

Exit Function

Nulo:
ConsumoIngredienteMes = 0

End Function
Function CostoConsumoMes(Mes As Integer, ano As Integer) As Double
'Costo Total de Ingredientes consumidos
'COSTO DE VENTA POR TIPO

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de costo de cantidades consumidas teoricos
miSQL = "SELECT DBIngredientes.TIPO1, Sum(IIf(ConsumoIngredienteMes([DBIngredientes]![CodIngrediente]," & Mes & "," & ano & ")=0,0,ConsumoIngredienteMes([DBIngredientes]![CodIngrediente]," & Mes & "," & ano & ")*CostoIngredienteGlobal([DBIngredientes]![CodIngrediente],numerosemana(ultimodomingomes(" & Mes & "," & ano & "))))) AS Total FROM DBIngredientes GROUP BY DBIngredientes.TIPO1 HAVING (((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoConsumoMes = rst("Total")
rst.Close

Exit Function

Nulo:
CostoConsumoMes = 0

End Function

    'METODO 2: Consumo = Inventario Final Pitaya Mes anterior + Compras (Ingresos ) - Inventario final pitaya del mes ( A nivel de ingrediente y luego agrupado por tipo de ingrediente)

    'METODO 3: Inventario Final mes anterior + COmpras - Inventario Final del mes ( a nivel total) da valor total
    '          Inventario Final Mes = InventarioCotizaionUltimasemana + InventarioIngredienteUltimaSemana + IngresosUltimosDIas - COnsumoUltimosDIas
               'CostoVariableMesGlobal = InventarioFinalMes(mes - 1, ano) + ComprasVariablesMes(mes, ano) - InventarioFinalMes(mes, ano)


Function ComprasVariablesMes(Mes As Integer, ano As Integer) As Double
'Monto gastado en compras de ingredientes VARIABLES  OK

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de Montos totales de costizaciones de ingredients variables
miSQL = "SELECT DBIngredientes.TIPO1, Month([IngresosPitaya]![Fecha]) AS Mes, Year([IngresosPitaya]![Fecha]) AS Ano, Sum([IngresosPitaya]![Cantidad]*CostoUnitarioUnidad([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS CostoTotal FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion GROUP BY DBIngredientes.TIPO1, Month([IngresosPitaya]![Fecha]), Year([IngresosPitaya]![Fecha]) HAVING (((DBIngredientes.TIPO1)='VARIABLES') AND ((Month([IngresosPitaya]![Fecha]))=" & Mes & ") AND ((Year([IngresosPitaya]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasVariablesMes = rst("CostoTotal")
rst.Close

Exit Function

Nulo:
ComprasVariablesMes = 0

End Function
Function InventarioFinalMes(Mes As Integer, ano As Integer) As Double
'Inventario final del mes usando formulas de abajo

InventarioFinalMes = CostoInventarioIngrediente(numerosemana(ultimodomingomes(Mes, ano))) + CostoInventarioCotizacion(numerosemana(ultimodomingomes(Mes, ano))) + CostoIngresosUltimosDias(Mes, ano) - CostoConsumoUltimosDias(Mes, ano)

End Function
               
Function CostoIngresosUltimosDias(Mes As Integer, ano As Integer) As Double
' Costo de los ingresos VARIABLES desde el inventario hasta el ultimo dia del mes

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim fei, fef As Date

fei = ultimodomingomes(Mes, ano) + 1
fef = DateSerial(ano, Mes, diasmes(DateSerial(ano, Mes, 1)))

'Hallar costo de ingresos entre dos  fechas especificas considerando las fechas
miSQL = "SELECT Sum([IngresosPitaya]![Cantidad]*CostoUnitarioUnidad([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS Costo, [IngresosPitaya]![Fecha]<=#" & fef & "# AS Expr1, [IngresosPitaya]![Fecha]>=#" & fei & "# AS Expr2, DBIngredientes.TIPO1 FROM DBIngredientes INNER JOIN (IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion) ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" GROUP BY [IngresosPitaya]![Fecha]<=#" & fef & "#, [IngresosPitaya]![Fecha]>=#" & fei & "#, DBIngredientes.TIPO1 HAVING ((([IngresosPitaya]![Fecha]<=#" & fef & "#)<>0) AND (([IngresosPitaya]![Fecha]>=#" & fei & "#)<>0) AND ((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoIngresosUltimosDias = rst("Costo")
rst.Close
     
Exit Function

Nulo:
CostoIngresosUltimosDias = 0

End Function

Function CantidadConsumoUltimosDias(Mes As Integer, ano As Integer, ING As String) As Double
' Cantidad de los Consumos VARIABLES desde el inventario hasta el ultimo dia del mes

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim fei, fef As Date

fei = ultimodomingomes(Mes, ano) + 1
fef = DateSerial(ano, Mes, diasmes(DateSerial(ano, Mes, 1)))

'Hallar cantidad  de CONSUMOS de cada ingrediente entre dos  fechas especificas considerando las fechas
miSQL = "SELECT [NotaDePedido]![Fecha]<=#" & fef & "# AS lim1," & _
" [NotaDePedido]![Fecha]>=#" & fei & "# AS lim2, SubReceta.CodIngrediente," & _
" Sum([SubReceta]![Cantidad]*[SubPedido]![Cantidad]*" & _
" FactorDeUsoMejorado([SubReceta]![CodIngrediente],[SubPedido]![Empaque],[NotaDePedido]![Modalidad],[NotaDePedido]![Fecha])) AS Consumo" & _
" FROM (NotaDePedido INNER JOIN SubPedido ON NotaDePedido.CodPedido = SubPedido.CodPedido)" & _
" INNER JOIN (DBBatidos INNER JOIN SubReceta ON DBBatidos.CodBatido = SubReceta.CodBatido)" & _
" ON SubPedido.CodBatido = DBBatidos.CodBatido" & _
" GROUP BY [NotaDePedido]![Fecha]<=#" & fef & "#, [NotaDePedido]![Fecha]>=#" & fei & "#, SubReceta.CodIngrediente" & _
" HAVING ((([NotaDePedido]![Fecha]<=#" & fef & "#)<>0) AND (([NotaDePedido]![Fecha]>=#" & fei & "#)<>0)" & _
" AND ((SubReceta.CodIngrediente)='" & ING & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CantidadConsumoUltimosDias = rst("Consumo")
rst.Close
     
Exit Function

Nulo:
CantidadConsumoUltimosDias = 0

End Function

Function CostoConsumoUltimosDias(Mes As Integer, ano As Integer) As Double
' Costo de los Consumos VARIABLES  (SIN PAN Y PAZ) desde el inventario hasta el ultimo dia del mes

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String


'Hallar costo  de CONSUMOS de cada ingrediente entre dos  fechas especificas considerando las fechas
miSQL = "SELECT DBIngredientes.TIPO1, Sum(IIf(CantidadConsumoUltimosDias(" & Mes & "," & ano & ",[DBIngredientes]![CodIngrediente])=0,0,CantidadConsumoUltimosDias(" & Mes & "," & ano & ",[DBIngredientes]![CodIngrediente])*CostoIngredienteGlobal([DBIngredientes]![CodIngrediente],numerosemana(ultimodomingomes(" & Mes & "," & ano & ")+1)))) AS Total FROM DBIngredientes GROUP BY DBIngredientes.TIPO1 HAVING (((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoConsumoUltimosDias = rst("Total")
rst.Close
     
Exit Function

Nulo:
CostoConsumoUltimosDias = 0

End Function

'GASTOS

'GASTOS OPERATIVOS Y ADMINISTRATIVOS Y CAPITAL

Function GastosMesIngreso(Tipo As String, Mes As Integer, ano As Integer) As Double
'Gastos de tipo de ingrediente especifico de un mes y ano dado
'COSTOS FIJOS CONTROLABLES

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de compras de un gast especifico segun tipo de un mes y ano dado
miSQL = "SELECT DBIngredientes.TIPO2, Month([IngresosPitaya]![Fecha]) AS Mes, Year([IngresosPitaya]![Fecha]) AS Ano, Sum([IngresosPitaya]![Cantidad]*CostoUnitarioUnidad([Cotizaciones]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS Monto FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN IngresosPitaya ON Cotizaciones.CodCotizacion = IngresosPitaya.CodCotizacion GROUP BY DBIngredientes.TIPO2, Month([IngresosPitaya]![Fecha]), Year([IngresosPitaya]![Fecha]) HAVING (((DBIngredientes.TIPO2)='" & Tipo & "') AND ((Month([IngresosPitaya]![Fecha]))=" & Mes & ") AND ((Year([IngresosPitaya]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
GastosMesIngreso = Round(rst("Monto"), 1)
rst.Close

Exit Function

Nulo:
GastosMesIngreso = 0

End Function

Function GastosMesFatura(Tipo As String, Mes As Integer, ano As Integer) As Double
'Gastos de tipo de ingrediente especifico de un mes y ano dado
'COSTOS FIJOS NO CONTROLABLES

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Suma de compras de un gast especifico segun tipo de un mes y ano dado
miSQL = "SELECT DBIngredientes.TIPO2, Month([Compras]![Fecha]) AS Mes, Year([Compras]![Fecha]) AS Ano, Sum([Compras]![CostoTotal]) AS Monto FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion GROUP BY DBIngredientes.TIPO2, Month([Compras]![Fecha]), Year([Compras]![Fecha]) HAVING (((DBIngredientes.TIPO2)='" & Tipo & "') AND ((Month([Compras]![Fecha]))=" & Mes & ") AND ((Year([Compras]![Fecha]))=" & ano & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
GastosMesFatura = Round(rst("Monto"), 1)
rst.Close

Exit Function

Nulo:
GastosMesFatura = 0

End Function

'INVENTARIO INICIAL MES - INVENTARIO FINAL MES (PRINCIPAL)
'INVENTARIO FINAL MES = INVENTARIO FINAL ULTIMO DOMINGO + COMPRAS EN ULTIMO LAPSO - INGRESOS A PITAYA 1 - INGRESOS PITAYA 2

Function InventarioMainFinalMes(Mes As Integer, ano As Integer) As Double
'Inventario final del mes usando formulas de abajo

InventarioMainFinalMes = CostoInventarioMain(numerosemana(ultimodomingomes(Mes, ano))) + ComprasUltimosDias(Mes, ano) - CostoIngresosMainUltimosDias(Mes, ano)

End Function
               
Function CostoIngresosMainUltimosDias(Mes As Integer, ano As Integer) As Double
' Costo de los ingresos VARIABLES desde el inventario hasta el ultimo dia del mes DEL ALMACEN PRINCIPAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim fei, fef As Date

fei = ultimodomingomes(Mes, ano) + 1
fef = DateSerial(ano, Mes, diasmes(DateSerial(ano, Mes, 1)))

'Hallar costo de ingresos entre dos  fechas especificas considerando las fechas
miSQL = "SELECT [IngresosPitaya]![Fecha]>=#" & fei & "# AS Desde, [IngresosPitaya]![Fecha]<=#" & fef & "# AS Hasta," & _
" Sum([IngresosPitaya]![Cantidad]*CostoUnitarioUnidad([IngresosPitaya]![CodCotizacion],numerosemana([IngresosPitaya]![Fecha]))) AS Total," & _
" DBIngredientes.TIPO1" & _
" FROM DBIngredientes" & _
" INNER JOIN (IngresosPitaya INNER JOIN Cotizaciones ON IngresosPitaya.CodCotizacion = Cotizaciones.CodCotizacion)" & _
" ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY [IngresosPitaya]![Fecha]>=#" & fei & "#, [IngresosPitaya]![Fecha]<=#" & fef & "#, DBIngredientes.TIPO1" & _
" HAVING ((([IngresosPitaya]![Fecha]>=#" & fei & "#)<>0) AND (([IngresosPitaya]![Fecha]<=#" & fef & "#)<>0)" & _
" AND ((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CostoIngresosMainUltimosDias = rst("Total")
rst.Close
     
Exit Function

Nulo:
CostoIngresosMainUltimosDias = 0

End Function

Function ComprasUltimosDias(Mes As Integer, ano As Integer) As Double
' Total de las compras VARIABLES desde el inventario hasta el ultimo dia del mes DEL ALMACEN PRINCIPAL

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim fei, fef As Date

fei = ultimodomingomes(Mes, ano) + 1
fef = DateSerial(ano, Mes, diasmes(DateSerial(ano, Mes, 1)))

'Hallar suma de cosnumo de compras entre dos  fechas especificas considerando las fechas
miSQL = "SELECT [Compras]![Fecha]>=#" & fei & "# AS Desde, [Compras]![Fecha]<=#" & fef & "# AS Hasta, Sum([Compras]![Cantidad]*CostoUnitarioUnidad([Compras]![CodCotizacion],numerosemana([Compras]![Fecha]))) AS Total, DBIngredientes.TIPO1 FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente) INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion GROUP BY [Compras]![Fecha]>=#" & fei & "#, [Compras]![Fecha]<=#" & fef & "#, DBIngredientes.TIPO1 HAVING ((([Compras]![Fecha]>=#" & fei & "#)<>0) AND (([Compras]![Fecha]<=#" & fef & "#)<>0) AND ((DBIngredientes.TIPO1)='VARIABLES'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComprasUltimosDias = rst("Total")
rst.Close
     
Exit Function

Nulo:
ComprasUltimosDias = 0

End Function

